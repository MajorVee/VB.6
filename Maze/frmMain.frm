VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Mistwalkers"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   13935
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   8715
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H00C0C0C0&
      Caption         =   "P L A Y"
      Height          =   735
      Left            =   11280
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Image cmdClose 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   13080
      Picture         =   "frmMain.frx":4C088
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Are you ready to take the Maze?"
      BeginProperty Font 
         Name            =   "Chiller"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   855
      Left            =   3120
      TabIndex        =   0
      Top             =   720
      Width           =   10335
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Options"
      Begin VB.Menu mnuMech 
         Caption         =   "Mechanics"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Mistwalkers") = vbYes Then
    End
End If
End Sub

Private Sub cmdPlay_Click()
Unload Me
frmLevel1.Show
End Sub

Private Sub mnuMech_Click()
frmMechanics.Show
End Sub

Private Sub mnuProg_Click()
frmMe.Show
End Sub
