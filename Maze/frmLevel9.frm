VERSION 5.00
Begin VB.Form frmLevel9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level 9"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLevel9.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmLevel9.frx":2512C
   ScaleHeight     =   7185
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   4200
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   4200
      Top             =   480
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   6720
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image imgPass 
      Height          =   495
      Left            =   5400
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   7
      Left            =   4200
      Picture         =   "frmLevel9.frx":2858E
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1575
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
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Image Image2 
      Height          =   855
      Index           =   6
      Left            =   6720
      Picture         =   "frmLevel9.frx":29AE2
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   855
      Index           =   5
      Left            =   4680
      Picture         =   "frmLevel9.frx":2B036
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   855
      Index           =   4
      Left            =   2520
      Picture         =   "frmLevel9.frx":2C58A
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   3
      Left            =   7920
      Picture         =   "frmLevel9.frx":2DADE
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   855
      Index           =   2
      Left            =   5760
      Picture         =   "frmLevel9.frx":2F032
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   975
      Index           =   1
      Left            =   4200
      Picture         =   "frmLevel9.frx":30586
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   855
      Index           =   0
      Left            =   2640
      Picture         =   "frmLevel9.frx":31ADA
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Image imgFin 
      Height          =   735
      Left            =   6480
      Picture         =   "frmLevel9.frx":3302E
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   13
      Left            =   6960
      Picture         =   "frmLevel9.frx":34051
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   12
      Left            =   6960
      Picture         =   "frmLevel9.frx":355A5
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   11
      Left            =   7320
      Picture         =   "frmLevel9.frx":36AF9
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   10
      Left            =   4680
      Picture         =   "frmLevel9.frx":3804D
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   9
      Left            =   4680
      Picture         =   "frmLevel9.frx":395A1
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   8
      Left            =   4680
      Picture         =   "frmLevel9.frx":3AAF5
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   7
      Left            =   4680
      Picture         =   "frmLevel9.frx":3C049
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   6
      Left            =   3120
      Picture         =   "frmLevel9.frx":3D59D
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   5
      Left            =   1560
      Picture         =   "frmLevel9.frx":3EAF1
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   4
      Left            =   0
      Picture         =   "frmLevel9.frx":40045
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   0
      Picture         =   "frmLevel9.frx":41599
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   0
      Picture         =   "frmLevel9.frx":42AED
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   0
      Picture         =   "frmLevel9.frx":44041
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   0
      Picture         =   "frmLevel9.frx":45595
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Image imgStart 
      Height          =   615
      Left            =   720
      Picture         =   "frmLevel9.frx":46AE9
      Stretch         =   -1  'True
      Top             =   6480
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
Attribute VB_Name = "frmLevel9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.imgFin.Enabled = False
Me.imgPass.Enabled = False
Me.lblMsg.Caption = "Move the head at the start button" & vbCrLf & "  to start the game..."
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
For j = 0 To Me.Image2.Count - 1
Me.Image2(j).Enabled = False
Next j
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation, "Mistwalkers"
Unload Me
frmLevel9.Show
End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation, "Mistwalkers"
Unload Me
frmLevel9.Show
End Sub

Private Sub imgPass_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgFin.Enabled = True
End Sub

Private Sub imgFin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgStart.Enabled = False
Me.Timer1.Enabled = False
Me.Timer2.Enabled = False
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
For j = 0 To Me.Image2.Count - 1
Me.Image2(j).Enabled = False
Next j
For i = 0 To Me.lblNext.Count - 1
Me.lblNext(i).Enabled = True
Next i
End Sub

Private Sub imgStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblMsg.Caption = "THE MOVING BRICKS!"
Me.imgPass.Enabled = True
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = True
Next i
For j = 0 To Me.Image2.Count - 1
Me.Image2(j).Enabled = True
Next j
End Sub

Private Sub lblNext_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
frmLevel10.Show
End Sub
Private Sub mnuQuit_Click()
If MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Mistwalkers") = vbYes Then
    End
End If
End Sub

Private Sub mnuRes_Click()
Unload Me
frmLevel1.Show
End Sub


Private Sub Timer1_Timer()
Me.Image2(0).Top = Me.Image2(0).Top + 100
If Me.Image2(0).Top >= 4200 Then
Me.Timer1.Enabled = False
Me.Timer2.Enabled = True
Else
End If

Me.Image2(2).Top = Me.Image2(2).Top + 100
If Me.Image2(2).Top >= 4200 Then
Me.Timer1.Enabled = False
Me.Timer2.Enabled = True
Else
End If
End Sub

Private Sub Timer2_Timer()
Me.Image2(0).Top = Me.Image2(0).Top - 100
If Me.Image2(0).Top <= 3480 Then
Me.Timer2.Enabled = False
Me.Timer1.Enabled = True
Else
End If

Me.Image2(2).Top = Me.Image2(2).Top - 100
If Me.Image2(2).Top <= 3480 Then
Me.Timer2.Enabled = False
Me.Timer1.Enabled = True
Else
End If
End Sub

