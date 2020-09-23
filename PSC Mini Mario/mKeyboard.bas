Attribute VB_Name = "mKeyboard"
'Declarations for Checking keyboard input
Public Const vbKeyAlt As Long = 18: Public Const vbKeyLeftAlt As Long = 164: Public Const vbKeyRightAlt As Long = 165
Public Const vbKeyLeftCtrl As Long = 162: Public Const vbKeyRightCtrl As Long = 163
Public Const vbKeyLeftShift As Long = 160: Public Const vbKeyRightShift As Long = 161
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Function IsKeyDown(ByVal Key As KeyCodeConstants) As Boolean
  IsKeyDown = GetKeyState(Key) < 0
End Function

