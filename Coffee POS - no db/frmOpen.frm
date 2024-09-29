VERSION 5.00
Begin VB.Form frmOpen 
   BackColor       =   &H00004080&
   Caption         =   "Kape-Libro Coffee Shop"
   ClientHeight    =   8475
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   Picture         =   "frmOpen.frx":0000
   ScaleHeight     =   8475
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H000000FF&
      Caption         =   "Close Shop"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Open Shop"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   2640
      Picture         =   "frmOpen.frx":1371D0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kape-Libro Coffee Shop"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   660
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   6075
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About the POS"
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFacts_Click()
frmKnows.Show
End Sub

Private Sub cmdClose_Click()
'Terminate the Program :(
If MsgBox("Are you sure you want to close the Shop?", vbQuestion + vbYesNo, "Kape-Libro Coffee Shop") = vbYes Then
    End
End If
End Sub

Private Sub cmdOpen_Click()
Unload Me
frmPOS.Show
End Sub

Private Sub mnuAbout_Click()
Dialog.Show
End Sub
