Attribute VB_Name = "mAudio"
'Wrapper for simple audio reproduction
Private Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallback As Long) As Long
Private ResM As String

Public Sub Start_Mp3(FileName$)
    ResM = Space(128)
   'Call mciSendString("set media time format milliseconds", ResM, 128, 0)
    Call mciSendString("open """ & FileName$ & """ alias media", ResM, 128, 0)
    Call mciSendString("play media", ResM, 128, 0)
End Sub

Public Sub ReStart_Mp3()
   ResM = Space(128)
   Call mciSendString("play media", ResM, 128, 0)
End Sub

Public Sub Stop_Mp3()
    ResM = Space(128)
    Call mciSendString("stop media", ResM, 128, 0)
End Sub

Public Sub Close_Mp3()
   ResM = Space(128)
   Call mciSendString("close media", ResM, 128, 0)
End Sub



