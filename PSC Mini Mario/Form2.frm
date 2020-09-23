VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Joystick Configuration"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4395
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Center"
      Height          =   1320
      Left            =   45
      TabIndex        =   16
      Top             =   1335
      Width           =   1200
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   960
         Left            =   120
         ScaleHeight     =   60
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   60
         TabIndex        =   17
         Top             =   225
         Width           =   960
         Begin VB.Shape Shape2 
            Height          =   570
            Left            =   165
            Shape           =   2  'Oval
            Top             =   180
            Width           =   570
         End
         Begin VB.Shape Shape1 
            Height          =   105
            Left            =   405
            Shape           =   2  'Oval
            Top             =   405
            Width           =   90
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Keys"
      Height          =   1260
      Left            =   1320
      TabIndex        =   5
      Top             =   1335
      Width           =   3015
      Begin VB.CommandButton Command1 
         Height          =   270
         Index           =   4
         Left            =   2280
         TabIndex        =   15
         Top             =   855
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Height          =   270
         Index           =   3
         Left            =   2280
         TabIndex        =   13
         Top             =   510
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Height          =   270
         Index           =   2
         Left            =   765
         TabIndex        =   11
         Top             =   780
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Height          =   270
         Index           =   1
         Left            =   2280
         TabIndex        =   9
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Height          =   270
         Index           =   0
         Left            =   750
         TabIndex        =   7
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Right"
         Height          =   210
         Index           =   4
         Left            =   1680
         TabIndex        =   14
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Left"
         Height          =   210
         Index           =   3
         Left            =   1680
         TabIndex        =   12
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Start"
         Height          =   210
         Index           =   2
         Left            =   165
         TabIndex        =   10
         Top             =   825
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Run"
         Height          =   210
         Index           =   1
         Left            =   1665
         TabIndex        =   8
         Top             =   225
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Jump"
         Height          =   210
         Index           =   0
         Left            =   165
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1260
      Left            =   1320
      TabIndex        =   1
      Top             =   30
      Width           =   3015
      Begin VB.OptionButton Option1 
         Caption         =   "Do not use  Analog Stick"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   4
         Top             =   270
         Width           =   2820
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Use Analog Stick for Walk and run"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   855
         Width           =   2820
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Use Analog Stick for Walk Only"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   570
         Width           =   2820
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   1170
      Left            =   90
      ScaleHeight     =   1110
      ScaleWidth      =   1095
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim KKeys(6) As Long
Private Sub Command1_Click(Index As Integer)
   Dim JoyInformation As JOYINFO
   Dim LStop As Boolean
   Do
     DoEvents
     joyGetPos JOYSTICK1, JoyInformation
     If JoyInformation.wButtons <> 0 Then
        Command1(Index).Caption = Hex(JoyInformation.wButtons)
        KKeys(Index) = JoyInformation.wButtons
        LStop = True
     End If
   Loop Until LStop
End Sub

Private Sub Form_Load()
   Initialize True
End Sub

Private Sub Picture2_Click()
   Dim JoyInformation As JOYINFO
   Dim LStop As Boolean
   Do
     DoEvents
     
     joyGetPos JOYSTICK1, JoyInformation
     Shape1.Move JoyInformation.wXpos / 1000 - 4, JoyInformation.wYpos / 1000 - 4
     If JoyInformation.wButtons <> 0 Then
        CenterX = JoyInformation.wXpos
        KKeys(5) = CenterX
        CenterY = JoyInformation.wYpos
        KKeys(6) = CenterY
        LStop = True
     End If
   Loop Until LStop
End Sub
