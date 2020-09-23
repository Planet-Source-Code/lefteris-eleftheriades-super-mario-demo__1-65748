VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enemies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Sorry, I didn't mean to add a class module to keep simlicity but it saves me time.
'This handles the drawing of an enemie to given coodinates on the screen
Private Type ImagePossition
   X As Integer
   Y As Integer
   W As Integer
   H As Integer
End Type

Dim Enemy(30) As ImagePossition

Private Sub SavePossition(EnemyFrame As ImagePossition, X%, Y%, W%, H%)
  EnemyFrame.X = X
  EnemyFrame.Y = Y
  EnemyFrame.W = W
  EnemyFrame.H = H
End Sub


Private Sub Class_Initialize()
  'Possitions of the graphics in the bitmap file (X,Y,Width,Height)
  'Ant
  SavePossition Enemy(0), 0, 0, 19, 15
  SavePossition Enemy(1), 21, 0, 19, 15
  SavePossition Enemy(2), 42, 0, 19, 15
  
  'Fish
  SavePossition Enemy(3), 0, 16, 20, 14
  SavePossition Enemy(4), 21, 16, 20, 14
  SavePossition Enemy(5), 41, 16, 20, 14
  SavePossition Enemy(6), 62, 16, 20, 14

  'Green Turtle
  SavePossition Enemy(7), 0, 31, 20, 23
  SavePossition Enemy(8), 21, 31, 20, 23
  SavePossition Enemy(9), 0, 55, 20, 23
  SavePossition Enemy(10), 21, 32, 20, 23
  SavePossition Enemy(11), 42, 32, 20, 23
  SavePossition Enemy(12), 63, 32, 20, 23
  SavePossition Enemy(13), 42, 32, 20, 23

  'Red Turtle
  SavePossition Enemy(14), 0, 79, 20, 23
  SavePossition Enemy(15), 21, 79, 20, 23
  SavePossition Enemy(16), 0, 102, 20, 23
  SavePossition Enemy(17), 21, 102, 20, 23
  SavePossition Enemy(18), 42, 79, 20, 23
  SavePossition Enemy(19), 63, 79, 20, 23
  SavePossition Enemy(20), 42, 102, 20, 23
  
  'hedjehok
  SavePossition Enemy(21), 0, 126, 20, 13
  SavePossition Enemy(22), 21, 126, 20, 13
  SavePossition Enemy(23), 42, 126, 20, 13
  SavePossition Enemy(24), 63, 126, 20, 13
  
  'Carnivore Plant Green
  SavePossition Enemy(25), 0, 139, 24, 17
  SavePossition Enemy(26), 24, 139, 24, 17
  
  'Carnivore Plant Red
  SavePossition Enemy(27), 0, 157, 24, 17
  SavePossition Enemy(28), 24, 157, 24, 17

  'Fireball
  SavePossition Enemy(29), 48, 139, 18, 13
  SavePossition Enemy(30), 66, 139, 18, 13
End Sub

Public Sub DrawEnemy(ByVal destDC&, ByVal X&, ByVal Y&, SrcDc&, ByVal Frame As Byte)
  BitBlt destDC, X, Y, Enemy(Frame).W, Enemy(Frame).H, SrcDc, Enemy(Frame).X + 83, Enemy(Frame).Y, SRCAND
  BitBlt destDC, X, Y, Enemy(Frame).W, Enemy(Frame).H, SrcDc, Enemy(Frame).X, Enemy(Frame).Y, SRCINVERT
End Sub

Public Property Get Width(ByVal Frame As Byte) As Integer
   Width = Enemy(Frame).W
End Property

Public Property Get Height(ByVal Frame As Byte) As Integer
   Height = Enemy(Frame).H
End Property

