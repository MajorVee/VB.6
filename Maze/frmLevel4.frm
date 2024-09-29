VERSION 5.00
Begin VB.Form frmLevel4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level 4"
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   9600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLevel4.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmLevel4.frx":2512C
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrClose 
      Interval        =   1
      Left            =   1560
      Top             =   6360
   End
   Begin VB.Timer tmrOpen 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   6360
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8520
      Top             =   3360
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8520
      Top             =   3840
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   19
      Left            =   120
      Picture         =   "frmLevel4.frx":2858E
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   18
      Left            =   1200
      Picture         =   "frmLevel4.frx":29AE2
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Image imgPass 
      Height          =   615
      Left            =   2640
      Top             =   4440
      Width           =   855
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
      Height          =   735
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Image imgFin 
      Height          =   735
      Left            =   240
      Picture         =   "frmLevel4.frx":2B036
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   17
      Left            =   120
      Picture         =   "frmLevel4.frx":2C059
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   16
      Left            =   5880
      Picture         =   "frmLevel4.frx":2D5AD
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   15
      Left            =   1800
      Picture         =   "frmLevel4.frx":2EB01
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   14
      Left            =   600
      Picture         =   "frmLevel4.frx":30055
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   13
      Left            =   7920
      Picture         =   "frmLevel4.frx":315A9
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   12
      Left            =   3960
      Picture         =   "frmLevel4.frx":32AFD
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   11
      Left            =   7560
      Picture         =   "frmLevel4.frx":34051
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   10
      Left            =   5640
      Picture         =   "frmLevel4.frx":355A5
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   9
      Left            =   3720
      Picture         =   "frmLevel4.frx":36AF9
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   8
      Left            =   3720
      Picture         =   "frmLevel4.frx":3804D
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   7
      Left            =   7560
      Picture         =   "frmLevel4.frx":395A1
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   6
      Left            =   120
      Picture         =   "frmLevel4.frx":3AAF5
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   5
      Left            =   2040
      Picture         =   "frmLevel4.frx":3C049
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   4
      Left            =   5040
      Picture         =   "frmLevel4.frx":3D59D
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   3
      Left            =   3120
      Picture         =   "frmLevel4.frx":3EAF1
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   2
      Left            =   600
      Picture         =   "frmLevel4.frx":40045
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   3480
      Picture         =   "frmLevel4.frx":41599
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   0
      Left            =   5640
      Picture         =   "frmLevel4.frx":42AED
      Stretch         =   -1  'True
      Top             =   960
      Width           =   2055
   End
   Begin VB.Image imgStart 
      Height          =   615
      Left            =   6240
      Picture         =   "frmLevel4.frx":44041
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuRes 
         Caption         =   "Restart this Level"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit Game"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmLevel4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.imgFin.Enabled = False
Me.imgPass.Enabled = False
Me.tmrClose.Enabled = False
Me.tmrOpen.Enabled = False
Me.lblMsg.Caption = "Move over the head to start button to start the game..."
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation, "Mistwalkers"
Me.tmrDown.Enabled = False
Me.tmrUp.Enabled = False
Unload Me
frmLevel4.Show
End Sub

Private Sub imgFin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgStart.Enabled = False
Me.tmrDown.Enabled = False
Me.tmrUp.Enabled = False
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

Private Sub imgStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblMsg.Caption = "Don't hit any blocks!!" & vbCrLf & "Be careful or your DEAD! HAHA :D"
Me.imgPass.Enabled = True

For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = True
Next i

Me.tmrDown.Enabled = True
Me.tmrClose.Enabled = True
End Sub

Private Sub lblNext_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
frmLevel5.Show
End Sub

Private Sub mnuQuit_Click()
If MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Mistwalkers") = vbYes Then
    End
End If
End Sub

Private Sub mnuRes_Click()
Unload Me
frmLevel4.Show
End Sub

Private Sub tmrClose_Timer()
Me.Image1(2).Left = Me.Image1(2).Left + 10.1

If Me.Image1(2).Left >= 1800 Then
Me.tmrClose.Enabled = False
Me.tmrOpen.Enabled = True
Else
End If

End Sub

Private Sub tmrDown_Timer()
Me.Image1(16).Top = Me.Image1(16).Top + 10.1

If Me.Image1(16).Top >= 3840 Then
Me.tmrDown.Enabled = False
Me.tmrUp.Enabled = True
Else
End If
End Sub

Private Sub tmrOpen_Timer()
Me.Image1(2).Left = Me.Image1(2).Left - 10.1

If Me.Image1(2).Left <= 600 Then
Me.tmrOpen.Enabled = False
Me.tmrClose.Enabled = True
Else
End If
End Sub

Private Sub tmrUp_Timer()
Me.Image1(16).Top = Me.Image1(16).Top - 10.1
If Me.Image1(16).Top <= 2400 Then
Me.tmrUp.Enabled = False
Me.tmrDown.Enabled = True
Else
End If
End Sub
