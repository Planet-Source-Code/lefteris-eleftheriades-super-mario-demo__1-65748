Attribute VB_Name = "mJoystick"
'Functions for checking and calibrating joystick input
'Thanks to Matt Carpenter- 2002/3
Option Explicit

Const MAXPNAMELEN = 32

Public Type JOYCAPS
        wMid As Integer
        wPid As Integer
        szPname As String * MAXPNAMELEN
        wXmin As Long
        wXmax As Long
        wYmin As Long
        wYmax As Long
        wZmin As Long
        wZmax As Long
        wNumButtons As Long
        wPeriodMin As Long
        wPeriodMax As Long
End Type

Public Type JOYINFO
        wXpos As Long
        wYpos As Long
        wZpos As Long
        wButtons As Long
End Type
Public Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Public Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long


'Joystick error codes and return values
Public Const JOYERR_NOERROR = 0
Public Const JOYERR_BASE As Long = 160
Public Const JOYERR_UNPLUGGED As Long = (JOYERR_BASE + 7)
Public Const MMSYSERR_BASE As Long = 0
Public Const MMSYSERR_NODRIVER As Long = (MMSYSERR_BASE + 6)
Public Const MMSYSERR_INVALPARAM As Long = (MMSYSERR_BASE + 11)
Public Const JOYERR_NOCANDO = (JOYERR_BASE + 6)
Public Const JOYERR_PARMS = (JOYERR_BASE + 5)

Public Const JOYSTICK1 As Long = &H0
Public Const JOYSTICK2 As Long = &H1
Public Const JOY_BUTTON2 = &H2
Public Const JOY_BUTTON1 = &H1
'This i added because i have an N64 controller :p (Tried to make N64 games in the past, didnt work)
Public Enum N64ControllerKeys
  CUp = 1
  CRight = 2
  CDown = 4
  CLeft = 8
  SideL = &H40
  SideR = &H80
  A = &H110
  B = &H220
  Z = &H400
  Start = &H800
  PadUp = &H1000
  PadRight = &H2000
  PadDown = &H4000
  PadLeft = &H8000
End Enum

Public Const N64PadLeft = &H8000
Public Const N64PadRight = &H2000
Public Const N64A = &H110
Public Const N64B = &H220
Public CenterX&, CenterY&, MaxX&, MaxY&, MinX&, MinY&

Public Function JoyStickCheck() As String
  'Returns: "OK " & Name, "Not Connected", "No Driver","Not Connected / Calibrated","Unknown Error","OK But Controller Not Connected with device adaptor"
  'According to the joystick status. If the function returns OK you can use joystick
  Dim rt As Long
  Dim JoyTestInfo As JOYINFO
  Dim JoyStickCaps As JOYCAPS
  rt = joyGetPos(JOYSTICK1, JoyTestInfo) 'See if there is a joystick
  If rt <> JOYERR_NOERROR Then
        If rt = JOYERR_UNPLUGGED Then
          JoyStickCheck = "Not Connected"
        ElseIf rt = MMSYSERR_NODRIVER Then
          JoyStickCheck = "No Driver"
        ElseIf rt = JOYERR_PARMS Then
          JoyStickCheck = "Not Connected / Calibrated"
        Else
          JoyStickCheck = "Unknown Error"
       End If
   Else
   If (JoyTestInfo.wButtons = 4095 Or (JoyTestInfo.wXpos = 0 And JoyTestInfo.wYpos = 0)) Then
      JoyStickCheck = "OK But Controller Not Connected with device adaptor"
   Else
      JoyStickCheck = "OK " & JoyStickCaps.szPname
   End If
End If
End Function

Public Sub Initialize(BoomDevice As Boolean)
'Note: BoomDevice is a controller to USB adaptor, that can be used instead of joystick
Dim rt As Long
Dim JoyTestInfo As JOYINFO
Dim JoyStickCaps As JOYCAPS
Dim SystemPath As String

rt = joyGetPos(JOYSTICK1, JoyTestInfo) 'See if there is a joystick

If rt <> JOYERR_NOERROR Then
    If rt = JOYERR_UNPLUGGED Then
        MsgBox "No controller connected", vbExclamation, "Error"
    ElseIf rt = MMSYSERR_NODRIVER Then
        MsgBox "No controller driver on system", vbExclamation, "Error"
    ElseIf rt = JOYERR_PARMS Then
        Shell "rundll32.exe shell32.dll,Control_RunDLL Joy.cpl ,%*", vbNormalFocus 'Show Gaming options control Panel
        
        MsgBox "Device not connected to USB -or- Controller badly calibrated." & vbCrLf & vbCrLf & "A) Connect the device and Run again." & vbCrLf & vbCrLf & "B) From the joypad control panel, check the controllers tab for " & vbCrLf & "    ""Monster Gamepad""" & vbCrLf & "     If it exists click properties, and see if in the test tab, the white box" & vbCrLf & "     has the cross on it's exact center and it moves when you move" & vbCrLf & "     the analogue stick." & vbCrLf & "     It not go to the ""Settings"" tab and click ""calibrate""", vbExclamation, "Error"
    Else
        MsgBox "Unknown Error", vbExclamation, "Error"
    End If
    Exit Sub
End If
If BoomDevice And (JoyTestInfo.wButtons = 4095 Or (JoyTestInfo.wXpos = 0 And JoyTestInfo.wYpos = 0)) Then
   Debug.Print "Controller Not Connected with device adaptor."
   Exit Sub
End If
  
  'Get the max and min position on the joystick
  joyGetDevCaps JOYSTICK1, JoyStickCaps, Len(JoyStickCaps)

  With JoyStickCaps
    MaxX = .wXmax
    MinX = .wXmin
    MaxY = .wYmax
    MinY = .wYmin
  End With
  CenterX = JoyTestInfo.wXpos
  CenterY = JoyTestInfo.wYpos
End Sub

Public Sub Calibrate()
   'Tell the user to center the stick and press a key on the joystick
   Dim JoyInformation As JOYINFO
   Do
     DoEvents
     joyGetPos JOYSTICK1, JoyInformation
     CenterX = JoyInformation.wXpos
     CenterY = JoyInformation.wYpos
   Loop Until JoyInformation.wButtons <> 0
End Sub

