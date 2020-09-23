VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   278
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   530
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pEnemies 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3045
      Left            =   8235
      Picture         =   "Editor.frx":0000
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   166
      TabIndex        =   10
      Top             =   390
      Visible         =   0   'False
      Width           =   2490
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   420
      Picture         =   "Editor.frx":18CBE
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   359
      TabIndex        =   9
      Top             =   675
      Width           =   5385
   End
   Begin VB.PictureBox Trees 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   8040
      Picture         =   "Editor.frx":1E160
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   8
      Top             =   810
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   6765
      MaskColor       =   &H00000000&
      Picture         =   "Editor.frx":2592A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   675
      Width           =   840
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   5865
      MaskColor       =   &H00000000&
      Picture         =   "Editor.frx":25FFC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   675
      Width           =   885
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8130
      Top             =   2745
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   15
      Max             =   190
      TabIndex        =   5
      Top             =   3885
      Width           =   7890
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   30
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   4
      Top             =   435
      Width           =   360
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   30
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   24
      TabIndex        =   3
      Top             =   45
      Width           =   360
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   2790
      Left            =   45
      ScaleHeight     =   182
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   520
      TabIndex        =   2
      Top             =   1080
      Width           =   7860
   End
   Begin VB.PictureBox Picture1 
      Height          =   690
      Left            =   405
      ScaleHeight     =   630
      ScaleWidth      =   7500
      TabIndex        =   0
      Top             =   0
      Width           =   7560
      Begin VB.PictureBox pTiles 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   0
         Picture         =   "Editor.frx":26F56
         ScaleHeight     =   84
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   507
         TabIndex        =   1
         Top             =   0
         Width           =   7605
      End
   End
   Begin VB.Menu mnuLevel 
      Caption         =   "Level"
      Begin VB.Menu LoadLevel 
         Caption         =   "Load"
      End
      Begin VB.Menu SaveLevel 
         Caption         =   "Save"
      End
      Begin VB.Menu ClearLevel 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SRCAND = &H8800C6
Private Const SRCINVERT = &H660046
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Dim Level(215, 12) As Byte
Dim Tool1&, Tool2&

Private Sub ClearLevel_Click()
  Dim X&, Y&
  For Y = 0 To 12
    For X = 0 To 215
      Level(X, Y) = 255
    Next X
  Next Y
End Sub

Private Sub Command1_Click(Index As Integer)
Tool1 = 25 * 3 + Index
End Sub

Private Sub Form_Load()
  ClearLevel_Click
End Sub

Private Sub HScroll1_Change()
  Draw
End Sub

Private Sub HScroll1_Scroll()
  Draw
End Sub

Private Sub LoadLevel_Click()
  CD1.Filter = "Super Mario Levels|*.sml|All Files|*.*"
  CD1.ShowOpen
  If CD1.FileName <> "" Then
  Open CD1.FileName For Binary As #1
    Get #1, , Level()
  Close #1
  Draw
  End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Picture2_MouseMove Button, 0, X, Y
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    Level(X \ 20 + HScroll1.Value, Y \ 14) = Tool1
    Draw
  End If
  If Button = 2 Then
    Level(X \ 20 + HScroll1.Value, Y \ 14) = Tool2
    Draw
  End If
  If Button = 4 Then
    Level(X \ 20 + HScroll1.Value, Y \ 14) = 255
    Draw
  End If
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    Picture3(0).Cls
    Picture3(0).PaintPicture Picture4.Picture, 0, 0, 20, 20, (X \ 20) * 20, 0, 20, 20
    Picture3(0).Refresh
    Tool1 = (X \ 20) + 77
    Me.Caption = Tool1 + 77
  End If
  If Button = 2 Then
    Picture3(1).Cls
    Picture3(1).PaintPicture Picture4.Picture, 0, 0, 20, 20, (X \ 20) * 20, 0, 20, 20
    Picture3(1).Refresh
    Tool2 = (X \ 20) + 77
  End If
End Sub

Private Sub pTiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  pTiles_MouseMove Button, 0, X, Y
End Sub

Sub Draw()
Picture2.Cls
For X = 0 To 25
  Picture2.Line (X * 20, 0)-(X * 20, 13 * 14), &HA0A0A0
Next
For Y = 0 To 12
  For X = 0 To 25
    Picture2.Line (0, (Y + 1) * 14)-(20 * 26, (Y + 1) * 14), &HA0A0A0
    If Level(X + HScroll1.Value, Y) < 25 * 3 Then
      BitBlt Picture2.hDC, X * 20, Y * 14, 20, 14, pTiles.hDC, (Level(X + HScroll1.Value, Y) Mod 25) * 20, 14 * (3 + (Level(X + HScroll1.Value, Y) \ 25)), SRCAND
      BitBlt Picture2.hDC, X * 20, Y * 14, 20, 14, pTiles.hDC, (Level(X + HScroll1.Value, Y) Mod 25) * 20, 14 * (Level(X + HScroll1.Value, Y) \ 25), SRCINVERT
    ElseIf Level(X + HScroll1.Value, Y) = 75 Then
       BitBlt Picture2.hDC, X * 20, Y * 14, 60, 28, Trees.hDC, 60, 0, SRCAND
       BitBlt Picture2.hDC, X * 20, Y * 14, 60, 28, Trees.hDC, 0, 0, SRCINVERT
    ElseIf Level(X + HScroll1.Value, Y) = 76 Then
       BitBlt Picture2.hDC, X * 20, Y * 14, 40, 14, Trees.hDC, 40, 57, SRCAND
       BitBlt Picture2.hDC, X * 20, Y * 14, 40, 14, Trees.hDC, 0, 57, SRCINVERT
    ElseIf Level(X + HScroll1.Value, Y) >= 77 Then
       Select Case Level(X + HScroll1.Value, Y)
         Case 77, 78
            BitBlt Picture2.hDC, X * 20, Y * 14, 21, 15, pEnemies.hDC, 83, 0, SRCAND
            BitBlt Picture2.hDC, X * 20, Y * 14, 21, 15, pEnemies.hDC, 0, 0, SRCINVERT
         Case 79
            BitBlt Picture2.hDC, X * 20, Y * 14, 21, 15, pEnemies.hDC, 83, 16, SRCAND
            BitBlt Picture2.hDC, X * 20, Y * 14, 21, 15, pEnemies.hDC, 0, 16, SRCINVERT
         Case 80
            BitBlt Picture2.hDC, X * 20, Y * 14 - 8, 21, 23, pEnemies.hDC, 83, 55, SRCAND
            BitBlt Picture2.hDC, X * 20, Y * 14 - 8, 21, 23, pEnemies.hDC, 0, 55, SRCINVERT
         Case 81
            BitBlt Picture2.hDC, X * 20, Y * 14 - 8, 21, 23, pEnemies.hDC, 83, 31, SRCAND
            BitBlt Picture2.hDC, X * 20, Y * 14 - 8, 21, 23, pEnemies.hDC, 0, 31, SRCINVERT
         Case 82
            BitBlt Picture2.hDC, X * 20, Y * 14 - 8, 21, 23, pEnemies.hDC, 83, 102, SRCAND
            BitBlt Picture2.hDC, X * 20, Y * 14 - 8, 21, 23, pEnemies.hDC, 0, 102, SRCINVERT
         Case 83
            BitBlt Picture2.hDC, X * 20, Y * 14 - 8, 21, 23, pEnemies.hDC, 83, 79, SRCAND
            BitBlt Picture2.hDC, X * 20, Y * 14 - 8, 21, 23, pEnemies.hDC, 0, 79, SRCINVERT
         Case 84
            BitBlt Picture2.hDC, X * 20, Y * 14, 21, 13, pEnemies.hDC, 125, 126, SRCAND
            BitBlt Picture2.hDC, X * 20, Y * 14, 21, 13, pEnemies.hDC, 42, 126, SRCINVERT
         Case 85
            BitBlt Picture2.hDC, X * 20, Y * 14, 21, 13, pEnemies.hDC, 83, 126, SRCAND
            BitBlt Picture2.hDC, X * 20, Y * 14, 21, 13, pEnemies.hDC, 0, 126, SRCINVERT
         Case 86
            BitBlt Picture2.hDC, X * 20 + 9, Y * 14 - 3, 24, 17, pEnemies.hDC, 83, 139, SRCAND
            BitBlt Picture2.hDC, X * 20 + 9, Y * 14 - 3, 24, 17, pEnemies.hDC, 0, 139, SRCINVERT
         Case 87
            BitBlt Picture2.hDC, X * 20 + 9, Y * 14 - 3, 24, 17, pEnemies.hDC, 83, 156, SRCAND
            BitBlt Picture2.hDC, X * 20 + 9, Y * 14 - 3, 24, 17, pEnemies.hDC, 0, 156, SRCINVERT
         Case 88
            BitBlt Picture2.hDC, X * 20, Y * 14, 20, 16, pEnemies.hDC, 83, 173, SRCAND
            BitBlt Picture2.hDC, X * 20, Y * 14, 20, 16, pEnemies.hDC, 0, 173, SRCINVERT
         Case 89, 90, 91
            BitBlt Picture2.hDC, X * 20, Y * 14 - 1, 20, 15, pEnemies.hDC, 83 + 20 * (Level(X + HScroll1.Value, Y) - 88), 173, SRCAND
            BitBlt Picture2.hDC, X * 20, Y * 14 - 1, 20, 15, pEnemies.hDC, 0 + 20 * (Level(X + HScroll1.Value, Y) - 88), 173, SRCINVERT
         Case 92
            BitBlt Picture2.hDC, X * 20, Y * 14 - 1, 24, 17, pEnemies.hDC, 131, 156, SRCAND
            BitBlt Picture2.hDC, X * 20, Y * 14 - 1, 24, 17, pEnemies.hDC, 48, 156, SRCINVERT
         Case 93
            BitBlt Picture2.hDC, X * 20 + 3, Y * 14 + 1, 14, 13, pEnemies.hDC, 83, 188, SRCAND
            BitBlt Picture2.hDC, X * 20 + 3, Y * 14 + 1, 14, 13, pEnemies.hDC, 0, 188, SRCINVERT
         Case 94
            BitBlt Picture2.hDC, X * 20 + 3, Y * 14 + 1, 14, 13, pEnemies.hDC, 125, 188, SRCAND
            BitBlt Picture2.hDC, X * 20 + 3, Y * 14 + 1, 14, 13, pEnemies.hDC, 42, 188, SRCINVERT
       End Select
    End If
  Next
Next
Picture2.Refresh
End Sub

Private Sub pTiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    Picture3(0).Cls
    Picture3(0).PaintPicture pTiles.Picture, 0, 0, 20, 14, (X \ 20) * 20, (Y \ 14) * 14, 20, 14
    Picture3(0).Refresh
    Tool1 = (X \ 20) + (Y \ 14) * 25
    Me.Caption = Tool1
  End If
  If Button = 2 Then
    Picture3(1).Cls
    Picture3(1).PaintPicture pTiles.Picture, 0, 0, 20, 14, (X \ 20) * 20, (Y \ 14) * 14, 20, 14
    Picture3(1).Refresh
    Tool2 = (X \ 20) + (Y \ 14) * 25
  End If
End Sub

Private Sub SaveLevel_Click()
  CD1.Filter = "Super Mario Levels|*.sml|All Files|*.*"
  CD1.ShowSave
  Open CD1.FileName For Binary As #1
    Put #1, , Level()
  Close #1
End Sub

