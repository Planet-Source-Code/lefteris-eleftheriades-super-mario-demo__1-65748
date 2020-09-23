Attribute VB_Name = "mWindow"
'Declarations for moving a widow that doesnt have a titlebar see Form_mousedown event. (can also be used for resizing)
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

