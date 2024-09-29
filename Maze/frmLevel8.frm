VERSION 5.00
Begin VB.Form frmLevel8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level 8"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLevel8.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmLevel8.frx":2512C
   ScaleHeight     =   7200
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtScore 
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtHide 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   600
      Top             =   1800
   End
   Begin VB.Image imgTimer 
      Height          =   735
      Left            =   8520
      Picture         =   "frmLevel8.frx":2858E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgPass 
      Height          =   615
      Left            =   2520
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   7320
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   5640
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   2640
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Image Image3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2880
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   4200
      Picture         =   "frmLevel8.frx":287AB
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image imgFin 
      Height          =   495
      Left            =   5880
      Picture         =   "frmLevel8.frx":29CFF
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image imgFinChange 
      Height          =   495
      Left            =   2520
      Picture         =   "frmLevel8.frx":2AD22
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   22
      Left            =   7440
      Picture         =   "frmLevel8.frx":2BD45
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   21
      Left            =   480
      Picture         =   "frmLevel8.frx":2D299
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   20
      Left            =   4080
      Picture         =   "frmLevel8.frx":2E7ED
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   19
      Left            =   7080
      Picture         =   "frmLevel8.frx":2FD41
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   18
      Left            =   120
      Picture         =   "frmLevel8.frx":31295
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   17
      Left            =   3840
      Picture         =   "frmLevel8.frx":327E9
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   16
      Left            =   7440
      Picture         =   "frmLevel8.frx":33D3D
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   15
      Left            =   4200
      Picture         =   "frmLevel8.frx":35291
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   14
      Left            =   480
      Picture         =   "frmLevel8.frx":367E5
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   13
      Left            =   7080
      Picture         =   "frmLevel8.frx":37D39
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   12
      Left            =   120
      Picture         =   "frmLevel8.frx":3928D
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   11
      Left            =   480
      Picture         =   "frmLevel8.frx":3A7E1
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   10
      Left            =   3840
      Picture         =   "frmLevel8.frx":3BD35
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   9
      Left            =   7440
      Picture         =   "frmLevel8.frx":3D289
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   8
      Left            =   7080
      Picture         =   "frmLevel8.frx":3E7DD
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   7
      Left            =   7560
      Picture         =   "frmLevel8.frx":3FD31
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   6
      Left            =   4200
      Picture         =   "frmLevel8.frx":41285
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   5
      Left            =   3840
      Picture         =   "frmLevel8.frx":427D9
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   4
      Left            =   120
      Picture         =   "frmLevel8.frx":43D2D
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   7200
      Picture         =   "frmLevel8.frx":45281
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   5400
      Picture         =   "frmLevel8.frx":467D5
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   3600
      Picture         =   "frmLevel8.frx":47D29
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   1800
      Picture         =   "frmLevel8.frx":4927D
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Image imgStart 
      Height          =   615
      Left            =   600
      Picture         =   "frmLevel8.frx":4A7D1
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   2520
      Picture         =   "frmLevel8.frx":4B26A
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Menu mnuO 
      Caption         =   "Options"
      Begin VB.Menu mnuRes 
         Caption         =   "Restart this Level"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit this Game"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmLevel8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgFin.Enabled = True
End Sub

Private Sub lblMsg_Click()
Me.txtScore.Text = Val(Me.txtScore.Text) + Val(Me.txtHide.Text)
frmLevel8.txtScore.Text = Me.txtScore.Text
frmLevel8.Show
Unload Me
End Sub

Private Sub mnuQuit_Click()
If MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Mistwalkers") = vbYes Then
    End
End If
End Sub

Private Sub mnuRes_Click()
Unload Me
frmLevel8.Show
End Sub
Dim a As Integer

Private Sub Form_Load()
a = 1
Me.imgFin.Enabled = False
Me.Image5.Enabled = False
Me.Timer1.Enabled = False
Me.lblMsg.Caption = "Move over the head to start button to start the game..."
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation, "Mistwalkers"
Unload Me
frmLevel8.Show
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation, "Mistwalkers"
Unload Me
frmLevel8.Show
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgFinChange.Visible = False
Me.Image4.Visible = True
Me.Image2.Visible = True
Me.Timer1.Enabled = False
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation, "Mistwalkers"
Unload Me
frmLevel8.Show
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgFin.Enabled = True
End Sub

Private Sub imgFin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgStart.Enabled = False
Me.Image3.Enabled = False
Me.Timer1.Enabled = False
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
For i = 0 To Me.lblNext.Count - 1
Me.lblNext(i).Enabled = True
Next i
End Sub

Private Sub imgStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblMsg.Caption = "Beware of choosing your way!" & vbCrLf & "Choose the right way without hitting the bricks."
Me.Image5.Enabled = True
Me.Image3.Enabled = True
Me.Timer1.Enabled = True
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = True
Next i

End Sub


Private Sub lblNext_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
frmLevel9.Show
End Sub

Private Sub Timer1_Timer()
On Error GoTo Err:
Me.imgTimer.Picture = LoadPicture(App.Path & "\" & a & ".gif")
a = a + 1
If a = 12 Then
MsgBox "Time is OVER! You're dead!", vbExclamation
Unload Me
End If
Exit Sub
Err:
a = 1
End Sub
