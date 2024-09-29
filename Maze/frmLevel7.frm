VERSION 5.00
Begin VB.Form frmLevel7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level 7"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLevel7.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmLevel7.frx":2512C
   ScaleHeight     =   7200
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   360
      Top             =   5520
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   4800
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   2400
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Image imgPass 
      Height          =   735
      Left            =   3000
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblBlock3 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label lblBlock2 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblBlock 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   7320
      TabIndex        =   1
      Top             =   2520
      Width           =   495
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
      Height          =   1335
      Left            =   5520
      TabIndex        =   0
      Top             =   5640
      Width           =   3735
   End
   Begin VB.Image imgFin 
      Height          =   615
      Left            =   2400
      Picture         =   "frmLevel7.frx":2858E
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   33
      Left            =   3960
      Picture         =   "frmLevel7.frx":295B1
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   32
      Left            =   3960
      Picture         =   "frmLevel7.frx":2AB05
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   31
      Left            =   3960
      Picture         =   "frmLevel7.frx":2C059
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   30
      Left            =   1320
      Picture         =   "frmLevel7.frx":2D5AD
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   29
      Left            =   3960
      Picture         =   "frmLevel7.frx":2EB01
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   28
      Left            =   1320
      Picture         =   "frmLevel7.frx":30055
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   27
      Left            =   120
      Picture         =   "frmLevel7.frx":315A9
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   26
      Left            =   1320
      Picture         =   "frmLevel7.frx":32AFD
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   25
      Left            =   3840
      Picture         =   "frmLevel7.frx":34051
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   24
      Left            =   3840
      Picture         =   "frmLevel7.frx":355A5
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   23
      Left            =   120
      Picture         =   "frmLevel7.frx":36AF9
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   22
      Left            =   1800
      Picture         =   "frmLevel7.frx":3804D
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   21
      Left            =   120
      Picture         =   "frmLevel7.frx":395A1
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   20
      Left            =   120
      Picture         =   "frmLevel7.frx":3AAF5
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   19
      Left            =   840
      Picture         =   "frmLevel7.frx":3C049
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   18
      Left            =   2040
      Picture         =   "frmLevel7.frx":3D59D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   17
      Left            =   3840
      Picture         =   "frmLevel7.frx":3EAF1
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   16
      Left            =   3240
      Picture         =   "frmLevel7.frx":40045
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   15
      Left            =   4440
      Picture         =   "frmLevel7.frx":41599
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   14
      Left            =   5640
      Picture         =   "frmLevel7.frx":42AED
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   13
      Left            =   4320
      Picture         =   "frmLevel7.frx":44041
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   12
      Left            =   6840
      Picture         =   "frmLevel7.frx":45595
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   855
      Index           =   11
      Left            =   5520
      Picture         =   "frmLevel7.frx":46AE9
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   10
      Left            =   8040
      Picture         =   "frmLevel7.frx":4803D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   9
      Left            =   8040
      Picture         =   "frmLevel7.frx":49591
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   8
      Left            =   8040
      Picture         =   "frmLevel7.frx":4AAE5
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   7
      Left            =   8040
      Picture         =   "frmLevel7.frx":4C039
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   6
      Left            =   5640
      Picture         =   "frmLevel7.frx":4D58D
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   5
      Left            =   5280
      Picture         =   "frmLevel7.frx":4EAE1
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   4
      Left            =   7680
      Picture         =   "frmLevel7.frx":50035
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   8040
      Picture         =   "frmLevel7.frx":51589
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   8040
      Picture         =   "frmLevel7.frx":52ADD
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   3960
      Picture         =   "frmLevel7.frx":54031
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   6840
      Picture         =   "frmLevel7.frx":55585
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Image imgStart 
      Height          =   615
      Left            =   4560
      Picture         =   "frmLevel7.frx":56AD9
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Menu mnuO 
      Caption         =   "Options"
      Begin VB.Menu mnuRes 
         Caption         =   "Restart this Level"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnnQuit 
         Caption         =   "Quit this Game"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmLevel7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuQuit_Click()
If MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Mistwalkers") = vbYes Then
    End
End If
End Sub

Private Sub mnuRes_Click()
Unload Me
frmLevel7.Show
End Sub

Private Sub Form_Load()
Me.imgFin.Enabled = False
Me.imgPass.Enabled = False
Me.lblMsg.Caption = "Move over the head to start button to start the game..."
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation
Unload Me
frmLevel7.Show
End Sub

Private Sub imgPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgFin.Enabled = True
End Sub

Private Sub imgFin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgStart.Enabled = False
Me.Timer1.Enabled = False
Me.Timer2.Enabled = False
Me.lblMsg.Caption = "LEVEL 3 COMPLETED!" & vbCrLf & "Click here to continue..."
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
For i = 0 To Me.lblNext.Count - 1
Me.lblNext(i).Enabled = True
Next i
End Sub

Private Sub imgStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblMsg.Caption = "Avoid the blinking blocks!" & vbCrLf & "You must pass without hitting the bricks."
Me.imgPass.Enabled = True
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = True
Next i


End Sub

Private Sub lblBlock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation
Unload Me
frmLevel7.Show
End Sub

Private Sub lblBlock2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation
Unload Me
frmLevel7.Show
End Sub

Private Sub lblBlock3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation
Unload Me
frmLevel7.Show
End Sub

Private Sub lblNext_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
frmLevel11.Show
End Sub

Private Sub Timer1_Timer()
Me.lblBlock.Visible = False
Me.lblBlock3.Visible = False
Me.lblBlock2.Visible = False
End Sub

Private Sub Timer2_Timer()
Me.lblBlock.Visible = True
Me.lblBlock3.Visible = True
Me.lblBlock2.Visible = True
End Sub
