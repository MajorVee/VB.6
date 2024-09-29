VERSION 5.00
Begin VB.Form frmLevel 
   BackColor       =   &H00000000&
   Caption         =   "Final Exam"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Take"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Final-Exam"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   3120
   End
   Begin VB.Image Image1 
      Height          =   1650
      Left            =   240
      Picture         =   "frmLevel.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1605
   End
End
Attribute VB_Name = "frmLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNo_Click()
'To show the Easy Form
Unload Me
frmExam.Show
End Sub

Private Sub cmdYes_Click()
'To show the Average Form
Unload Me
frmYes.Show
End Sub

Private Sub imgEnd_Click()
'To terminate the Program
If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Final Exam") = vbYes Then
    End
End If
End Sub
