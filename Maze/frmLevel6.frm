VERSION 5.00
Begin VB.Form frmLevel6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level 6"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLevel6.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmLevel6.frx":2512C
   ScaleHeight     =   7185
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgPass 
      Height          =   375
      Left            =   7320
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   1440
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label lblE 
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label lblD 
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7440
      TabIndex        =   4
      Top             =   5040
      Width           =   255
   End
   Begin VB.Label lblC 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblB 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   7080
      TabIndex        =   2
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label lblA 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   840
      Width           =   255
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
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Image imgFin 
      Height          =   735
      Left            =   1440
      Picture         =   "frmLevel6.frx":2858E
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   48
      Left            =   240
      Picture         =   "frmLevel6.frx":295B1
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   47
      Left            =   240
      Picture         =   "frmLevel6.frx":2AB05
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   46
      Left            =   600
      Picture         =   "frmLevel6.frx":2C059
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   45
      Left            =   2400
      Picture         =   "frmLevel6.frx":2D5AD
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   44
      Left            =   2040
      Picture         =   "frmLevel6.frx":2EB01
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   43
      Left            =   3480
      Picture         =   "frmLevel6.frx":30055
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   42
      Left            =   4920
      Picture         =   "frmLevel6.frx":315A9
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   41
      Left            =   6360
      Picture         =   "frmLevel6.frx":32AFD
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   40
      Left            =   7800
      Picture         =   "frmLevel6.frx":34051
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   39
      Left            =   7800
      Picture         =   "frmLevel6.frx":355A5
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   38
      Left            =   7800
      Picture         =   "frmLevel6.frx":36AF9
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   37
      Left            =   5880
      Picture         =   "frmLevel6.frx":3804D
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   36
      Left            =   5880
      Picture         =   "frmLevel6.frx":395A1
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   35
      Left            =   5880
      Picture         =   "frmLevel6.frx":3AAF5
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   34
      Left            =   7800
      Picture         =   "frmLevel6.frx":3C049
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   33
      Left            =   7800
      Picture         =   "frmLevel6.frx":3D59D
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   32
      Left            =   3840
      Picture         =   "frmLevel6.frx":3EAF1
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   31
      Left            =   3840
      Picture         =   "frmLevel6.frx":40045
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   30
      Left            =   3840
      Picture         =   "frmLevel6.frx":41599
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   29
      Left            =   3840
      Picture         =   "frmLevel6.frx":42AED
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   28
      Left            =   3600
      Picture         =   "frmLevel6.frx":44041
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   27
      Left            =   3360
      Picture         =   "frmLevel6.frx":45595
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   26
      Left            =   5640
      Picture         =   "frmLevel6.frx":46AE9
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   25
      Left            =   7800
      Picture         =   "frmLevel6.frx":4803D
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   24
      Left            =   7080
      Picture         =   "frmLevel6.frx":49591
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   23
      Left            =   5880
      Picture         =   "frmLevel6.frx":4AAE5
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   22
      Left            =   7800
      Picture         =   "frmLevel6.frx":4C039
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   21
      Left            =   7800
      Picture         =   "frmLevel6.frx":4D58D
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   20
      Left            =   7800
      Picture         =   "frmLevel6.frx":4EAE1
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   19
      Left            =   5880
      Picture         =   "frmLevel6.frx":50035
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   18
      Left            =   5880
      Picture         =   "frmLevel6.frx":51589
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   17
      Left            =   7800
      Picture         =   "frmLevel6.frx":52ADD
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   16
      Left            =   5880
      Picture         =   "frmLevel6.frx":54031
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   15
      Left            =   3600
      Picture         =   "frmLevel6.frx":55585
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   14
      Left            =   3840
      Picture         =   "frmLevel6.frx":56AD9
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   13
      Left            =   7800
      Picture         =   "frmLevel6.frx":5802D
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   12
      Left            =   6360
      Picture         =   "frmLevel6.frx":59581
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   11
      Left            =   7800
      Picture         =   "frmLevel6.frx":5AAD5
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   10
      Left            =   0
      Picture         =   "frmLevel6.frx":5C029
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   9
      Left            =   5760
      Picture         =   "frmLevel6.frx":5D57D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   8
      Left            =   7200
      Picture         =   "frmLevel6.frx":5EAD1
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   7
      Left            =   5760
      Picture         =   "frmLevel6.frx":60025
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   6
      Left            =   3840
      Picture         =   "frmLevel6.frx":61579
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   5
      Left            =   2880
      Picture         =   "frmLevel6.frx":62ACD
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   4
      Left            =   1440
      Picture         =   "frmLevel6.frx":64021
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   3
      Left            =   4320
      Picture         =   "frmLevel6.frx":65575
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   2
      Left            =   4320
      Picture         =   "frmLevel6.frx":66AC9
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   1
      Left            =   2880
      Picture         =   "frmLevel6.frx":6801D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   1440
      Picture         =   "frmLevel6.frx":69571
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image imgStart 
      Height          =   735
      Left            =   120
      Picture         =   "frmLevel6.frx":6AAC5
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1335
   End
   Begin VB.Menu munO 
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
Attribute VB_Name = "frmLevel6"
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
Me.lblA.Enabled = False
Me.lblB.Enabled = False
Me.lblC.Enabled = False
Me.lblD.Enabled = False
Me.lblE.Enabled = False
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation, "Mistwalkers"
Unload Me
frmLevel6.Show
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

Private Sub imgStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblMsg.Caption = "Get the letters alphabetically without hitting the bricks!"
Me.imgPass.Enabled = True
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = True
Next i
Me.lblA.Enabled = True
If Me.lblA.Visible = True Then
Me.imgFin.Enabled = False
Else
Me.imgFin.Enabled = True
End If

End Sub

Private Sub lblA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblA.Visible = False
Me.lblA.Enabled = False
Me.lblB.Enabled = True
End Sub

Private Sub lblB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblB.Visible = False
Me.lblB.Enabled = False
Me.lblC.Enabled = True
End Sub

Private Sub lblC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblC.Visible = False
Me.lblC.Enabled = False
Me.lblD.Enabled = True
End Sub

Private Sub lblD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblD.Visible = False
Me.lblD.Enabled = False
Me.lblE.Enabled = True
End Sub

Private Sub lblE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblE.Visible = False
Me.lblE.Enabled = False
Me.imgFin.Enabled = True
End Sub

Private Sub lblNext_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
frmLevel7.Show
End Sub

Private Sub mnuQuit_Click()
If MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Mistwalkers") = vbYes Then
    End
End If
End Sub

Private Sub mnuRes_Click()
Unload Me
frmLevel6.Show
End Sub
