Attribute VB_Name = "mResolution"
'This code is for changing the screen resolution I pray it can get smaller
'Borrowed this from API Guide.
Option Explicit
Const WM_DISPLAYCHANGE = &H7E
Const HWND_BROADCAST = &HFFFF&
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Const CCDEVICENAME = 32
Const CCFORMNAME = 32
Const DM_BITSPERPEL = &H40000
Const DM_PELSWIDTH = &H80000
Const DM_PELSHEIGHT = &H100000
Const CDS_UPDATEREGISTRY = &H1
Const CDS_TEST As Long = &H2
Const CDS_FULLSCREEN = &H4
Const DISP_CHANGE_SUCCESSFUL = 0
Const DISP_CHANGE_RESTART = 1
Const BITSPIXEL = 12
Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
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
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public OldX As Long, OldY As Long
Public nDC As Long

Public Sub TaskBar(ByVal Visible As Boolean)
    ShowWindow FindWindow("Shell_TrayWnd", vbNullString), -Visible * 5
    'FindWindowex
End Sub

Public Sub AlwaysOnTop(FrmHandle As Long, OnTop As Boolean)
    SetWindowPos FrmHandle, -OnTop - 2, 0, 0, 0, 0, 3
End Sub

Public Sub ChangeRes(X As Long, Y As Long, Bits As Long, Permanent As Boolean)
    Dim DevM As DEVMODE, ScInfo As Long, erg As Long, an As VbMsgBoxResult
    'Get the info into DevM
    erg = EnumDisplaySettings(0&, 0&, DevM)
    'This is what we're going to change
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = X 'ScreenWidth
    DevM.dmPelsHeight = Y 'ScreenHeight
    DevM.dmBitsPerPel = Bits '(can be 8, 16, 24, 32 or even 4)
    'Now change the display and check if possible
    erg = ChangeDisplaySettings(DevM, CDS_FULLSCREEN)
    'Check if succesfull
    Select Case erg&
        Case DISP_CHANGE_RESTART
           ' an = MsgBox("You've to reboot", vbYesNo + vbSystemModal, "Info")
           ' If an = vbYes Then
           '     erg& = ExitWindowsEx(EWX_REBOOT, 0&)
           ' End If
        Case DISP_CHANGE_SUCCESSFUL
            erg = ChangeDisplaySettings(DevM, Permanent + 2)
            ScInfo = Y * 2 ^ 16 + X ' &H(YY)(XX)
            'Notify all the windows of the screen resolution change
            SendMessage HWND_BROADCAST, WM_DISPLAYCHANGE, ByVal Bits, ByVal ScInfo
            'MsgBox "Everything's ok", vbOKOnly + vbSystemModal, "It worked!"
        Case Else
            'MsgBox "Mode not supported", vbOKOnly + vbSystemModal, "Error"
    End Select
End Sub
Public Sub SetResolution()
    Dim nDC As Long
    'retrieve the screen's resolution
    OldX = Screen.Width / Screen.TwipsPerPixelX
    OldY = Screen.Height / Screen.TwipsPerPixelY
    'Create a device context, compatible with the screen
    nDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
    'Change the screen's resolution
    ChangeRes 320, 200, GetDeviceCaps(nDC, BITSPIXEL), False
    '512x384
    '480x360
    '400x300
    '320x200
End Sub
Sub Main()
    'Emergency Resolution Restore function
    Dim nDC As Long
    If Screen.Width / Screen.TwipsPerPixelX = 320 Then
      'Create a device context, compatible with the screen
      nDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
      'Change the screen's resolution
      ChangeRes 800, 600, GetDeviceCaps(nDC, BITSPIXEL), False
      MsgBox "Emergency resolution restored after an unnormal termination of the game!"
    End If
    DoEvents
    DeleteDC nDC
    TaskBar True
    ShowCursor True
    DoEvents
    End
End Sub
Public Sub UnsetRes()
    'restore the screen resolution
    If OldX <> 320 Then
       ChangeRes OldX, OldY, GetDeviceCaps(nDC, BITSPIXEL), False
    Else
       ChangeRes 800, 600, GetDeviceCaps(nDC, BITSPIXEL), False
       MsgBox "Resolution Resored to 800x600"
    End If
    'delete our device context
    DeleteDC nDC
End Sub
