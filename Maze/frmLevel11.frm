VERSION 5.00
Begin VB.Form frmLevel11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level8"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9555
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLevel11.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmLevel11.frx":2512C
   ScaleHeight     =   7170
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtScore 
      Height          =   495
      Left            =   8760
      TabIndex        =   1
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtHide 
      Height          =   285
      Left            =   8760
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   200
      Top             =   6480
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   2400
      Picture         =   "frmLevel11.frx":2858E
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgStart 
      Height          =   495
      Left            =   120
      Picture         =   "frmLevel11.frx":28AB0
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   1320
      Picture         =   "frmLevel11.frx":29549
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   3120
      Picture         =   "frmLevel11.frx":2AA9D
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   4920
      Picture         =   "frmLevel11.frx":2BFF1
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   6720
      Picture         =   "frmLevel11.frx":2D545
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   4
      Left            =   120
      Picture         =   "frmLevel11.frx":2EA99
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   5
      Left            =   4320
      Picture         =   "frmLevel11.frx":2FFED
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   6
      Left            =   3720
      Picture         =   "frmLevel11.frx":31541
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   7
      Left            =   7080
      Picture         =   "frmLevel11.frx":32A95
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   8
      Left            =   7560
      Picture         =   "frmLevel11.frx":33FE9
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   9
      Left            =   7200
      Picture         =   "frmLevel11.frx":3553D
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   10
      Left            =   4320
      Picture         =   "frmLevel11.frx":36A91
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   11
      Left            =   360
      Picture         =   "frmLevel11.frx":37FE5
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   12
      Left            =   120
      Picture         =   "frmLevel11.frx":39539
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   13
      Left            =   7560
      Picture         =   "frmLevel11.frx":3AA8D
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   14
      Left            =   360
      Picture         =   "frmLevel11.frx":3BFE1
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   15
      Left            =   3720
      Picture         =   "frmLevel11.frx":3D535
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   16
      Left            =   7200
      Picture         =   "frmLevel11.frx":3EA89
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   17
      Left            =   4320
      Picture         =   "frmLevel11.frx":3FFDD
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   18
      Left            =   120
      Picture         =   "frmLevel11.frx":41531
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   19
      Left            =   7560
      Picture         =   "frmLevel11.frx":42A85
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   20
      Left            =   3720
      Picture         =   "frmLevel11.frx":43FD9
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   21
      Left            =   360
      Picture         =   "frmLevel11.frx":4552D
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   22
      Left            =   7200
      Picture         =   "frmLevel11.frx":46A81
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Image imgFinChange 
      Height          =   495
      Left            =   2400
      Picture         =   "frmLevel11.frx":47FD5
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image imgFin 
      Height          =   495
      Left            =   6000
      Picture         =   "frmLevel11.frx":48FF8
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   3720
      Picture         =   "frmLevel11.frx":4A01B
      Stretch         =   -1  'True
      Top             =   5760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image3 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      Top             =   1560
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   2640
      Top             =   5880
      Width           =   615
   End
   Begin VB.Image imgTimer 
      Height          =   735
      Left            =   8400
      Picture         =   "frmLevel11.frx":4B56F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmLevel11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub Form_Load()
a = 1
Me.txtHide.Text = 1
Me.imgFin.Enabled = False
Me.Image5.Enabled = False
Me.Timer1.Enabled = False
Me.lblMsg.Caption = "Point the start arrow" & vbCrLf & "  to start the game..."
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation
Unload Me
frmLevel11.Show
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation
Unload Me
frmLevel11.Show
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgFinChange.Visible = False
Me.Image4.Visible = True
Me.Image2.Visible = True
Me.Timer1.Enabled = False

End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation
Unload Me
frmLevel11.Show
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgFin.Enabled = True
End Sub

Private Sub imgFin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgStart.Enabled = False
Me.Image3.Enabled = False
Me.Timer1.Enabled = False
Me.lblMsg.Caption = "LEVEL  8 COMPLETED!" & vbCrLf & "Click here to continue..."
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
End Sub

Private Sub imgStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblMsg.Caption = "Choose the right way without hitting the bricks."
Me.Image5.Enabled = True
Me.Image3.Enabled = True
Me.Timer1.Enabled = True
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = True
Next i
End Sub

Private Sub lblMsg_Click()
Me.txtScore.Text = Val(Me.txtScore.Text) + Val(Me.txtHide.Text)
frmLevel11.txtScore.Text = Me.txtScore.Text
frmLevel11.Show
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
frmLevel11.Show
End If
Exit Sub
Err:
a = 1
End Sub

