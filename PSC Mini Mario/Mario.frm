VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Super Mario Mini (Paused)"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Mario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Water 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFAA55&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   15
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   6
      Top             =   2805
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Trees 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   3090
      Picture         =   "Mario.frx":0E42
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   4
      Top             =   1620
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.PictureBox pTiles 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   60
      Picture         =   "Mario.frx":860C
      ScaleHeight     =   84
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   507
      TabIndex        =   2
      Top             =   2310
      Visible         =   0   'False
      Width           =   7605
   End
   Begin VB.PictureBox pScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2730
      Left            =   0
      Picture         =   "Mario.frx":27A5E
      ScaleHeight     =   182
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   165
      Width           =   4800
      Begin VB.PictureBox Arch 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -15
         Picture         =   "Mario.frx":52520
         ScaleHeight     =   24
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   211
         TabIndex        =   5
         Top             =   1845
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.PictureBox pMario 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2430
         Left            =   3705
         Picture         =   "Mario.frx":56102
         ScaleHeight     =   162
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   120
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.PictureBox pEnemies 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3045
         Left            =   2685
         Picture         =   "Mario.frx":64514
         ScaleHeight     =   203
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   166
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   2490
         Begin VB.PictureBox pRagnaros 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1740
            Left            =   90
            Picture         =   "Mario.frx":7D1D2
            ScaleHeight     =   116
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   76
            TabIndex        =   7
            Top             =   615
            Visible         =   0   'False
            Width           =   1140
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Original Game and Graphics by Mike Wiering (Turbo Pascal)
'2.5 Hours of programming to make this beautiful game demo
'Modules are standalone and can be used freely
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Const TileWidth& = 20
Const TileHeight& = 14
Const MarioWidth& = 20
Const MarioHeight& = 27

Private Type EnemyAI
  X As Long
  Y As Integer
  oX As Currency
  dX As Currency
  F As Byte
  Dead As Boolean
End Type

Dim Level(215, 12) As Byte
Dim ExitFlag As Boolean
Dim BlockX&, OffsetX&, MarioX&, MarioY As Long
Dim dX&, dY&, MarioFrame%, dvy#, MarioFace As Byte
Dim TreeMovement As Boolean, DoneActionForThisJump As Boolean
Dim Enemy As New Enemies
Dim numEnemies As Integer
Dim EnemysInFieldX() As EnemyAI
Dim Jumping As Boolean
Dim Ragnaros As Boolean
Private Sub Form_Load()

'Open App.Path & "\test level.sml" For Binary As #1
  Open App.Path & "\enemy.sml" For Binary As #1
      Get #1, , Level
  Close #1
  MarioX = 70 + 30
  MarioY = 14 * 10 - MarioHeight - 1
  
  'Me.Move 0, 0
  Me.Show
  'SetResolution
  PlayGame
  'UnsetRes
  End
End Sub

Sub PlayGame()
Dim PreviewsTick As Long
  Do
    If GetTickCount >= PreviewsTick + 30 Then
       PreviewsTick = GetTickCount
       CheckKeyboardInput
       CollisionDetection
       
       MarioY = MarioY + dY
       OffsetX = OffsetX + dX
       If OffsetX >= TileWidth Then
          OffsetX = OffsetX - TileWidth
          BlockX = BlockX + 1
       End If
       If OffsetX < 0 Then
          OffsetX = OffsetX + TileWidth
          BlockX = BlockX - 1
       End If
       
       pScreen.Cls
       DrawTiles
       If Command = "-d" Then DrawGrid
       DrawEnemies
       DrawMario
       If MarioY > TileHeight * 13 Then
         'Mario fell off the stage
         MarioX = 70 + 30
         MarioY = 14 * 10 - MarioHeight - 1
         BlockX = 0
         OffsetX = 0
                Dim j
                For j = 1 To numEnemies
                   EnemysInFieldX(j).Dead = False
                Next j
       End If
       If Command = "-d" Then DrawDebug
       pScreen.Refresh
    End If
    DoEvents
  Loop Until ExitFlag
End Sub

Sub CheckKeyboardInput()
    Static MarioWalkTicker As Byte, TreeAnimTicker As Byte
    dX = 0
    'dY = Round(IIf(Abs(dvy) > 1, dvy, 1 * Sgn(dvy)))
    dY = dvy
    dvy = dvy + 0.3
    If dvy > 4 Then dvy = 4
    If IsKeyDown(vbKeyEscape) Then ExitFlag = True
    If IsKeyDown(vbKeyLeft) Then
       dX = -3
       MarioFace = 1
    End If
    If IsKeyDown(vbKeyRight) Then
       dX = 3
       MarioFace = 0
    End If
    'If IsKeyDown(vbKeyUp) Then dY = -2
    'If IsKeyDown(vbKeyDown) Then dY = 2
    If IsKeyDown(vbKeyAlt) And NotInAir Then
       dvy = -6
       If IsKeyDown(vbKeyControl) Then dvy = -7
       DoneActionForThisJump = False
       Jumping = True
    End If
    
    If IsKeyDown(vbKeyControl) Then
       dX = dX * 2
    End If
    
    'Screen Scroll
    If IsKeyDown(vbKeyLeftShift) Then
       MarioX = MarioX + 2
       dX = dX - 2
    End If
    If IsKeyDown(vbKeyRightShift) Then
       MarioX = MarioX - 2
       dX = dX + 2
    End If
    
    'Walk Animation
    MarioWalkTicker = MarioWalkTicker + 1
    If MarioWalkTicker = 3 Then
       If dX <> 0 Then MarioFrame = 1 - MarioFrame
       MarioWalkTicker = 0
    End If
    
    TreeAnimTicker = TreeAnimTicker + 1
    If TreeAnimTicker = 10 Then
       TreeMovement = Not TreeMovement
       TreeAnimTicker = 0
    End If
End Sub

Sub CheckBrick(X&, Y&, Ordinate As Byte)
 'Used by CollisionDetection routine
 'Check for collision with a specific brick
 If Solid(Level(X, Y)) = True Then
    'Collision on the Y-Ordinate
    If Ordinate = 1 Then
       dY = dY - Sgn(dY)
       If dY = 0 Then
          dvy = 0
          Jumping = False
       End If
    End If
    'Collision on the X-Ordinate
    If Ordinate = 0 Then dX = dX - Sgn(dX)
    
    'If a collision was found, reduce movement step by 1 (eg if i told it to move 3 pixels, now i tell it to move 2) and check again if move can be made
    CollisionDetection
 End If
End Sub

Function CheckBrickHitAction(X&, Y&) As Boolean
   If Level(X, Y) = 25 Then
      Level(X, Y) = 255
      dvy = 0
      CheckBrickHitAction = True
   End If
   If Level(X, Y) = 28 Then
      Level(X, Y) = 29
      dvy = 0
      CheckBrickHitAction = True
   End If
End Function

Sub CollisionDetection()
  Dim MoveCanNotBeMade As Boolean
  On Error Resume Next
  'Action detection
  If dY < 0 Then
    'Only one hit action is allowed per jump with priority the brick mario is under more than the other
    If Not DoneActionForThisJump Then
      If (OffsetX + MarioX) Mod TileWidth <= 10 Then
       DoneActionForThisJump = CheckBrickHitAction(BlockX + (OffsetX + MarioX) \ TileWidth, (MarioY + dY) \ TileHeight)
       If Not DoneActionForThisJump Then DoneActionForThisJump = CheckBrickHitAction(BlockX + (OffsetX + MarioX + dX + MarioWidth - 1) \ TileWidth, (MarioY + dY) \ TileHeight)
      Else
       DoneActionForThisJump = CheckBrickHitAction(BlockX + (OffsetX + MarioX + dX + MarioWidth - 1) \ TileWidth, (MarioY + dY) \ TileHeight)
       If Not DoneActionForThisJump Then DoneActionForThisJump = CheckBrickHitAction(BlockX + (OffsetX + MarioX) \ TileWidth, (MarioY + dY) \ TileHeight)
      End If
    End If
  End If
  
  'X-Ordinate
    If dX <> 0 Then
      CheckBrick BlockX + (OffsetX + MarioX + dX) \ TileWidth, (MarioY + MarioHeight) \ TileHeight, 0
      CheckBrick BlockX + (OffsetX + MarioX + MarioWidth + dX - 1) \ TileWidth, (MarioY + MarioHeight) \ TileHeight, 0
      CheckBrick BlockX + (OffsetX + MarioX + dX) \ TileWidth, (MarioY + MarioHeight / 2) \ TileHeight, 0
      CheckBrick BlockX + (OffsetX + MarioX + MarioWidth + dX - 1) \ TileWidth, (MarioY + MarioHeight / 2) \ TileHeight, 0
      CheckBrick BlockX + (OffsetX + MarioX + dX) \ TileWidth, MarioY \ TileHeight, 0
      CheckBrick BlockX + (OffsetX + MarioX + MarioWidth + dX - 1) \ TileWidth, MarioY \ TileHeight, 0
    End If
  
  'Y-Ordinate (but taking to account the dx change of the X Ordinate collision detection)
    If dY <> 0 Then
      CheckBrick BlockX + (OffsetX + MarioX + dX) \ TileWidth, (MarioY + MarioHeight + dY) \ TileHeight, 1
      CheckBrick BlockX + (OffsetX + MarioX + dX + MarioWidth - 1) \ TileWidth, (MarioY + MarioHeight + dY) \ TileHeight, 1
      CheckBrick BlockX + (OffsetX + MarioX + dX) \ TileWidth, (MarioY + dY) \ TileHeight, 1
      CheckBrick BlockX + (OffsetX + MarioX + dX + MarioWidth - 1) \ TileWidth, (MarioY + dY) \ TileHeight, 1
    End If
End Sub

Function NotInAir() As Boolean
  Dim Flag As Boolean, BlockY&
  NotInAir = False
  BlockY = (MarioY + MarioHeight + 3) \ TileHeight
  If BlockY >= 0 And BlockY <= 12 Then
    If BlockX + (OffsetX + MarioX) \ TileWidth >= 0 And BlockX + (OffsetX + MarioX + MarioWidth - 1) \ TileWidth <= 215 Then
      If Solid(Level(BlockX + (OffsetX + MarioX) \ TileWidth, BlockY)) Then Flag = True
      If Solid(Level(BlockX + (OffsetX + MarioX + MarioWidth - 1) \ TileWidth, BlockY)) Then Flag = True
      NotInAir = Flag
    End If
  End If
End Function

Sub DrawTiles()
  Dim X&, Y&, i&
  Static WaterStep&
  'Handle Water animation
  WaterStep = WaterStep + 1
  If WaterStep > 14 Then WaterStep = 0
  BitBlt Water.hdc, 0, WaterStep - 14, 20, 14, pTiles.hdc, 21 * 20, 0, SRCCOPY
  BitBlt Water.hdc, 0, WaterStep, 20, 14, pTiles.hdc, 21 * 20, 0, SRCCOPY
  
  'Draw Arc
  For i = 0 To 2
    BitBlt pScreen.hdc, -((BlockX * 20 + OffsetX) / 2 Mod 211) + 211 * i, 125, 211, 12, Arch.hdc, 0, 12, SRCAND
    BitBlt pScreen.hdc, -((BlockX * 20 + OffsetX) / 2 Mod 211) + 211 * i, 125, 211, 12, Arch.hdc, 0, 0, SRCINVERT
  Next
  
  For Y = 0 To 12
    For X = -2 To 16
      If BlockX + X >= 0 And BlockX + X <= 159 Then
         If Level(BlockX + X, Y) <= 76 Then
           If Level(BlockX + X, Y) = 21 Then
             'Water
             BitBlt pScreen.hdc, X * TileWidth - OffsetX, Y * TileHeight, TileWidth, TileHeight, Water.hdc, 0, 0, SRCCOPY
           ElseIf Level(BlockX + X, Y) = 20 Then
             'Water
             BitBlt pScreen.hdc, X * TileWidth - OffsetX, Y * TileHeight + 5, TileWidth, TileHeight, Water.hdc, 0, 5, SRCCOPY
           ElseIf Level(BlockX + X, Y) = 75 Then
             'Tree
             BitBlt pScreen.hdc, X * TileWidth - OffsetX, Y * TileHeight, 60, 28, Trees.hdc, 60, -28 * TreeMovement, SRCAND
             BitBlt pScreen.hdc, X * TileWidth - OffsetX, Y * TileHeight, 60, 28, Trees.hdc, 0, -28 * TreeMovement, SRCINVERT
           ElseIf Level(BlockX + X, Y) = 76 Then
             'Weed
             BitBlt pScreen.hdc, X * TileWidth - OffsetX, Y * TileHeight, 40, 14, Trees.hdc, 40, 57 - 14 * TreeMovement, SRCAND
             BitBlt pScreen.hdc, X * TileWidth - OffsetX, Y * TileHeight, 40, 14, Trees.hdc, 0, 57 - 14 * TreeMovement, SRCINVERT
           Else
             'All terrain
             BitBlt pScreen.hdc, X * TileWidth - OffsetX, Y * TileHeight, TileWidth, TileHeight, pTiles.hdc, (Level(BlockX + X, Y) Mod 25) * TileWidth, (Level(BlockX + X, Y) \ 25) * TileHeight + TileHeight * 3, SRCAND
             BitBlt pScreen.hdc, X * TileWidth - OffsetX, Y * TileHeight, TileWidth, TileHeight, pTiles.hdc, (Level(BlockX + X, Y) Mod 25) * TileWidth, (Level(BlockX + X, Y) \ 25) * TileHeight, SRCINVERT
           End If
         End If
      
      End If
    Next X
  Next Y
End Sub

Sub DrawGrid()
  Dim X&, Y&
  For Y = 0 To 12
      pScreen.Line (0, Y * TileHeight)-(16 * TileWidth, Y * TileHeight), &HFFAA00
  Next Y
  For X = 0 To 16
      pScreen.Line (X * TileWidth - OffsetX, 0)-(X * TileWidth - OffsetX, TileHeight * 13), &HFFAA00
      pScreen.CurrentX = X * TileWidth - OffsetX + (20 - pScreen.TextWidth(X + BlockX)) / 2 - 3
      pScreen.CurrentY = 0
      pScreen.Print X + BlockX
      'pScreen.Line (X * TileWidth, 0)-(X * TileWidth, TileHeight * 13), &HFFAA00
  Next X
End Sub

Sub DrawMario()
  BitBlt pScreen.hdc, MarioX, MarioY + 2, MarioWidth, MarioHeight, pMario.hdc, 60 + MarioWidth * MarioFrame, MarioHeight + MarioHeight * 3 * MarioFace, SRCAND
  BitBlt pScreen.hdc, MarioX, MarioY + 2, MarioWidth, MarioHeight, pMario.hdc, 0 + MarioWidth * MarioFrame, MarioHeight + MarioHeight * 3 * MarioFace, SRCINVERT
  'pScreen.Line (MarioX, MarioY)-(MarioX + 19, MarioY + 27), , B
End Sub

Sub DrawDebug()
  pScreen.CurrentY = 0
  pScreen.CurrentX = 0
  'pScreen.Print OffsetX
  'pScreen.Print BlockX
  'pScreen.Print MarioY
  
End Sub

Sub DrawEnemies()
    Dim X%, Y%
    For Y = 0 To 12
      For X = 0 To 16
         If BlockX + X >= 0 And BlockX + X <= 159 Then
            If Level(BlockX + X, Y) > 76 And Level(BlockX + X, Y) < 96 Then
              AddEnemy BlockX + X, Y, 0 '25
              Level(BlockX + X, Y) = 255
            End If
         End If
      Next
    Next
    Static EnemyStepTick As Integer
    Static EnemyStep As Boolean
    EnemyStepTick = EnemyStepTick + 1
    If EnemyStepTick >= 4 Then
       EnemyStepTick = 0
       EnemyStep = Not EnemyStep
    End If
    Dim i&
On Error Resume Next
    pScreen.CurrentX = 0
    pScreen.CurrentY = 0
    For i = 1 To numEnemies
      If EnemysInFieldX(i).Dead = False Then
        EnemysInFieldX(i).oX = EnemysInFieldX(i).oX + EnemysInFieldX(i).dX

        'Enemy Collision Detection
        Dim s&
        s = ((EnemysInFieldX(i).X - BlockX) * 20 - EnemysInFieldX(i).oX) \ 20
        If Command = "-d" Then pScreen.Line (s * 20 - OffsetX, 14 * EnemysInFieldX(i).Y)-(s * 20 + 20 - OffsetX, 14 * (EnemysInFieldX(i).Y + 1)), IIf(Solid(Level(s + BlockX, EnemysInFieldX(i).Y)), vbRed, vbGreen), B
        If Solid(Level(s + BlockX, EnemysInFieldX(i).Y)) Then EnemysInFieldX(i).dX = -1
        s = ((EnemysInFieldX(i).X - BlockX) * 20 - EnemysInFieldX(i).oX + Enemy.Width(EnemysInFieldX(i).F)) \ 20
        If Command = "-d" Then pScreen.Line (s * 20 - OffsetX, 14 * EnemysInFieldX(i).Y)-(s * 20 + 20 - OffsetX, 14 * (EnemysInFieldX(i).Y + 1)), IIf(Solid(Level(s + BlockX, EnemysInFieldX(i).Y)), vbRed, vbGreen), B
        If Solid(Level(s + BlockX, EnemysInFieldX(i).Y)) Then EnemysInFieldX(i).dX = 1
   
        If MarioX + BlockX * 20 + OffsetX > EnemysInFieldX(i).X * 20 - EnemysInFieldX(i).oX - 20 And MarioX + BlockX * 20 + OffsetX < EnemysInFieldX(i).X * 20 - EnemysInFieldX(i).oX + 20 Then
           'MarioX and EnemyX are the same
           If MarioY + MarioHeight > (EnemysInFieldX(i).Y + 1) * 14 - Enemy.Height(EnemysInFieldX(i).F) Then
             If dvy > 0 And Jumping Then
               'You kill an enemy
                EnemysInFieldX(i).Dead = True
                dvy = -dvy * 1.5
             Else
               'Enemy Kills you
                
                MarioX = 100
                MarioY = 14 * 10 - MarioHeight - 1
                BlockX = 0
                OffsetX = 0
                Dim j
                For j = 1 To numEnemies
                   EnemysInFieldX(j).Dead = False
                Next j
             End If
           End If
        End If
        Enemy.DrawEnemy pScreen.hdc, (EnemysInFieldX(i).X - BlockX) * 20 - OffsetX - EnemysInFieldX(i).oX, (EnemysInFieldX(i).Y + 1) * 14 - Enemy.Height(EnemysInFieldX(i).F), pEnemies.hdc, EnemysInFieldX(i).F + Abs(EnemyStep)
      End If
    Next i
End Sub

Sub AddEnemy(ByVal X&, ByVal Y&, F&)
  numEnemies = numEnemies + 1
  ReDim Preserve EnemysInFieldX(numEnemies)
  EnemysInFieldX(numEnemies).X = X
  EnemysInFieldX(numEnemies).Y = Y
  EnemysInFieldX(numEnemies).dX = 1 '0.5
  EnemysInFieldX(numEnemies).F = F
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ReleaseCapture
  SendMessage Me.hwnd, 161, 2, 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ExitFlag = True
End Sub

Function Solid(Index As Byte) As Boolean
   Select Case Index
      Case 20, 21, 40 To 44, 47, 48, 255, 75, 76, 77 To 94: Solid = False
      Case Else: Solid = True
   End Select
End Function
