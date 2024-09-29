VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000007&
   Caption         =   "Form3"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14115
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   7950
   ScaleWidth      =   14115
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
      Height          =   735
      Left            =   240
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   3135
   End
   Begin VB.CommandButton cmdModerate 
      BackColor       =   &H00FF8080&
      Caption         =   "Moderate"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10440
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   4125
      Left            =   10440
      Picture         =   "Form3.frx":1CBE8
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   3390
   End
   Begin VB.Image Image1 
      Height          =   2880
      Left            =   720
      Picture         =   "Form3.frx":21CD3
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   4560
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bring it on!"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1110
      Left            =   6720
      TabIndex        =   3
      Top             =   2640
      Width           =   4995
   End
   Begin VB.Image imgEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   13320
      Picture         =   "Form3.frx":2AF94
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You think you have the Guts?!"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1110
      Left            =   3240
      TabIndex        =   0
      Top             =   1440
      Width           =   10515
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
'To show Form 1
Unload Me
Form1.Show
End Sub

Private Sub cmdModerate_Click()
Unload Me
Form4.Show
End Sub


Private Sub imgEnd_Click()
'To end the Program
If MsgBox("Are you sure you want to stop playing?", vbQuestion + vbYesNo, "System") = vbYes Then
    End
End If
End Sub
