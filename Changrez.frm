VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Change Resolutions"
   ClientHeight    =   4665
   ClientLeft      =   2280
   ClientTop       =   2610
   ClientWidth     =   4575
   Height          =   5070
   Left            =   2220
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   4575
   Top             =   2265
   Width           =   4695
   Begin VB.CommandButton Command1 
      Caption         =   "Switch!"
      Height          =   495
      Left            =   3300
      TabIndex        =   2
      Top             =   1980
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Available Resolutions"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

' ChangRez sample by Matt Hart - mhart@taascforce.com
' http://www.webczar.com/defcon/mh
'
' How to change video resolution in Windows 95.
' The WinSDK API declarations for VB does NOT include
' this useful procedure, nor does it include some
' of the needed constants.  I had to figure out the API
' declaration and go to the Platform SDK
' find most of the constants.

Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const DM_DISPLAYFLAGS = &H200000
Const DM_DISPLAYFREQUENCY = &H400000

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpInitData As DEVMODE, ByVal dwFlags As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (lpszDeviceName As Any, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Const BITSPIXEL = 12

' /* Flags for ChangeDisplaySettings */
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST = &H2
Const CDS_FULLSCREEN = &H4
Const CDS_GLOBAL = &H8
Const CDS_SET_PRIMARY = &H10
Const CDS_RESET = &H40000000
Const CDS_SETRECT = &H20000000
Const CDS_NORESET = &H10000000

' /* Return values for ChangeDisplaySettings */
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const DISP_CHANGE_FAILED = -1
Const DISP_CHANGE_BADMODE = -2
Const DISP_CHANGE_NOTUPDATED = -3
Const DISP_CHANGE_BADFLAGS = -4
Const DISP_CHANGE_BADPARAM = -5

Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4

Dim D() As DEVMODE, lNumModes As Long

Private Sub Command1_Click()
    Dim l As Long, Flags As Long, x As Long
    x = List1.ListIndex
    D(x).dmFields = DM_BITSPERPEL Or DM_PELSWIDTH Or DM_PELSHEIGHT
    Flags = CDS_UPDATEREGISTRY
    l = ChangeDisplaySettings(D(x), Flags)
    Select Case l
        Case DISP_CHANGE_RESTART
            l = MsgBox("This change will not take effect until you reboot the system.  Reboot now?", vbYesNo)
            If l = vbYes Then
                Flags = 0
                l = ExitWindowsEx(EWX_REBOOT, Flags)
            End If
        Case DISP_CHANGE_SUCCESSFUL
        Case Else
            MsgBox "Error changing resolution! Returned: " & l
    End Select
End Sub

Private Sub Form_Load()
    Dim l As Long, lMaxModes As Long
    Dim lBits As Long, lWidth As Long, lHeight As Long
    lBits = GetDeviceCaps(hdc, BITSPIXEL)
    lWidth = Screen.Width \ Screen.TwipsPerPixelX
    lHeight = Screen.Height \ Screen.TwipsPerPixelY
    lMaxModes = 8
    ReDim D(0 To lMaxModes) As DEVMODE
    lNumModes = 0
    l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    Do While l
        List1.AddItem D(lNumModes).dmPelsWidth & "x" & D(lNumModes).dmPelsHeight & "x" & D(lNumModes).dmBitsPerPel
        If lBits = D(lNumModes).dmBitsPerPel And _
           lWidth = D(lNumModes).dmPelsWidth And _
           lHeight = D(lNumModes).dmPelsHeight Then
            List1.ListIndex = List1.NewIndex
        End If
        lNumModes = lNumModes + 1
        If lNumModes > lMaxModes Then
            lMaxModes = lMaxModes + 8
            ReDim Preserve D(0 To lMaxModes) As DEVMODE
        End If
        l = EnumDisplaySettings(ByVal 0, lNumModes, D(lNumModes))
    Loop
    lNumModes = lNumModes - 1
End Sub
