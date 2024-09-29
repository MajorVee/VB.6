VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   9495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13575
   LinkTopic       =   "Form6"
   ScaleHeight     =   9495
   ScaleWidth      =   13575
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H8000000A&
      Caption         =   "<<Back to Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   3615
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "(Don't mind the background...Trip ko lang gamitin :D)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   8520
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form6.frx":0000
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Instructions:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   285
      TabIndex        =   0
      Top             =   2160
      Width           =   1875
   End
   Begin VB.Image Image1 
      Height          =   9255
      Left            =   0
      Picture         =   "Form6.frx":0103
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15135
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Options"
      Begin VB.Menu mnuEasy 
         Caption         =   "Back to Easy Game"
      End
      Begin VB.Menu mnuMode 
         Caption         =   "Back to Moderate Game"
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
Form1.Show
End Sub

Private Sub mnuEasy_Click()
Unload Me
Form2.Show
End Sub

Private Sub mnuMode_Click()
Unload Me
Form4.Show
End Sub
