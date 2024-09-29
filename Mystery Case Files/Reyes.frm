VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   9720
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   14745
   LinkTopic       =   "Form1"
   Picture         =   "Reyes.frx":0000
   ScaleHeight     =   9720
   ScaleWidth      =   14745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Hard"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   10560
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Pick the level of your Game :D"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   3360
      TabIndex        =   2
      Top             =   960
      Width           =   10455
   End
   Begin VB.Image imgEnd 
      Height          =   735
      Left            =   480
      Picture         =   "Reyes.frx":BEE3D
      Stretch         =   -1  'True
      Top             =   360
      Width           =   735
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuInfo 
         Caption         =   "How to Play the Game"
      End
      Begin VB.Menu mnuWho 
         Caption         =   "Who Programmed it?"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHow_Click()
'To show Form6/Instructions
Unload Me
Form6.Show
End Sub

Private Sub Command1_Click()
'To show Form 2
Unload Me
Form2.Show
End Sub

Private Sub Command2_Click()
'To show Form 3
Unload Me
Form3.Show
End Sub

Private Sub imgEnd_Click()
'End the Program
If MsgBox("Are you sure you don't want to Play? TnT", vbQuestion + vbYesNo, "System") = vbYes Then
    End
End If
End Sub

Private Sub mnuInfo_Click()
Unload Me
Form6.Show
End Sub

Private Sub mnuWho_Click()
Unload Me
Form5.Show
End Sub
