VERSION 5.00
Begin VB.Form frmLevel5 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level 5"
   ClientHeight    =   7185
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLevel5.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmLevel5.frx":2512C
   ScaleHeight     =   7185
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgSec 
      Height          =   495
      Left            =   1920
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   1155
      Left            =   1890
      TabIndex        =   2
      Top             =   2040
      Width           =   4155
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image imgPass 
      Height          =   495
      Left            =   5400
      Top             =   3600
      Width           =   615
   End
   Begin VB.Image imgFin 
      Height          =   735
      Left            =   360
      Picture         =   "frmLevel5.frx":2858E
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   30
      Left            =   120
      Picture         =   "frmLevel5.frx":295B1
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   29
      Left            =   2040
      Picture         =   "frmLevel5.frx":2AB05
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   28
      Left            =   3960
      Picture         =   "frmLevel5.frx":2C059
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   27
      Left            =   6720
      Picture         =   "frmLevel5.frx":2D5AD
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   26
      Left            =   7680
      Picture         =   "frmLevel5.frx":2EB01
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   25
      Left            =   5880
      Picture         =   "frmLevel5.frx":30055
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   24
      Left            =   4320
      Picture         =   "frmLevel5.frx":315A9
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   23
      Left            =   3960
      Picture         =   "frmLevel5.frx":32AFD
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   22
      Left            =   2040
      Picture         =   "frmLevel5.frx":34051
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   21
      Left            =   120
      Picture         =   "frmLevel5.frx":355A5
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   20
      Left            =   120
      Picture         =   "frmLevel5.frx":36AF9
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   19
      Left            =   3840
      Picture         =   "frmLevel5.frx":3804D
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   18
      Left            =   120
      Picture         =   "frmLevel5.frx":395A1
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   17
      Left            =   5760
      Picture         =   "frmLevel5.frx":3AAF5
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   16
      Left            =   2040
      Picture         =   "frmLevel5.frx":3C049
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   15
      Left            =   3960
      Picture         =   "frmLevel5.frx":3D59D
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   14
      Left            =   1920
      Picture         =   "frmLevel5.frx":3EAF1
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   13
      Left            =   7560
      Picture         =   "frmLevel5.frx":40045
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   12
      Left            =   2040
      Picture         =   "frmLevel5.frx":41599
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   11
      Left            =   0
      Picture         =   "frmLevel5.frx":42AED
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   10
      Left            =   5640
      Picture         =   "frmLevel5.frx":44041
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   9
      Left            =   3720
      Picture         =   "frmLevel5.frx":45595
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   8
      Left            =   3960
      Picture         =   "frmLevel5.frx":46AE9
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   7
      Left            =   5880
      Picture         =   "frmLevel5.frx":4803D
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   6
      Left            =   0
      Picture         =   "frmLevel5.frx":49591
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   5
      Left            =   2040
      Picture         =   "frmLevel5.frx":4AAE5
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   4
      Left            =   5760
      Picture         =   "frmLevel5.frx":4C039
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   3
      Left            =   1920
      Picture         =   "frmLevel5.frx":4D58D
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   2
      Left            =   3840
      Picture         =   "frmLevel5.frx":4EAE1
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   120
      Picture         =   "frmLevel5.frx":50035
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   0
      Left            =   5760
      Picture         =   "frmLevel5.frx":51589
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Image imgStart 
      Height          =   615
      Left            =   240
      Picture         =   "frmLevel5.frx":52ADD
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuRes 
         Caption         =   "Restart"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit Game"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmLevel5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.imgFin.Enabled = False
Me.imgPass.Enabled = False
Me.lblMsg.Caption = "Move over the head to start button to start the game..."
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation, "Mistwalkers"
Unload Me
frmLevel5.Show
End Sub

Private Sub imgFin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgStart.Enabled = False
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
For i = 0 To Me.lblNext.Count - 1
Me.lblNext(i).Enabled = True
Next i
End Sub

Private Sub imgPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgFin.Enabled = True
End Sub

Private Sub imgSec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Image1(14).Visible = False
End Sub

Private Sub imgStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Me.lblMsg.Caption = "Don't hit any blocks!!" & vbCrLf & "Be careful or your DEAD! HAHA :D"
Me.imgPass.Enabled = True
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = True
Next i

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.tmrClose.Enabled = True
Me.tmrOpen.Enabled = True
End Sub

Private Sub lblNext_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
frmLevel6.Show
End Sub

Private Sub mnuQuit_Click()
If MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Mistwalkers") = vbYes Then
    End
End If
End Sub

Private Sub mnuRes_Click()
Unload Me
frmLevel5.Show
End Sub
