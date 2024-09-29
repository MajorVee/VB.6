VERSION 5.00
Begin VB.Form frmLevel10 
   BackColor       =   &H00004080&
   Caption         =   "Level 10"
   ClientHeight    =   8520
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11910
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmLevel10.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1680
      Top             =   4920
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1680
      Top             =   5640
   End
   Begin VB.Timer tmrUp2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   1440
   End
   Begin VB.Timer tmrDown2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   1920
   End
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   1920
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5400
      Top             =   1440
   End
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6240
      Top             =   4920
   End
   Begin VB.Timer tmrOpen 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7200
      Top             =   4920
   End
   Begin VB.Label lblBlock 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   2400
      TabIndex        =   9
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblBlock2 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   6240
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblBlock3 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   5520
      Width           =   855
   End
   Begin VB.Label lblE 
      Alignment       =   2  'Center
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
      Left            =   6720
      TabIndex        =   6
      Top             =   7320
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
      Left            =   2880
      TabIndex        =   5
      Top             =   7320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblC 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   3480
      Width           =   615
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
      Left            =   4440
      TabIndex        =   3
      Top             =   1680
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
      Left            =   8040
      TabIndex        =   2
      Top             =   1680
      Width           =   255
   End
   Begin VB.Image imgSec2 
      Height          =   495
      Left            =   6120
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   44
      Left            =   5880
      Picture         =   "frmLevel10.frx":2512C
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   43
      Left            =   3960
      Picture         =   "frmLevel10.frx":26680
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Image imgStart 
      Height          =   615
      Left            =   9360
      Picture         =   "frmLevel10.frx":27BD4
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Image imgPass 
      Height          =   495
      Left            =   8400
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image imgSec 
      Height          =   495
      Left            =   360
      Top             =   7800
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   42
      Left            =   3960
      Picture         =   "frmLevel10.frx":2866D
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   6120
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   6960
      TabIndex        =   0
      Top             =   3360
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   6600
      MouseIcon       =   "frmLevel10.frx":29BC1
      Picture         =   "frmLevel10.frx":4ECED
      Top             =   4080
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   41
      Left            =   9720
      Picture         =   "frmLevel10.frx":C6BC4
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   40
      Left            =   9720
      Picture         =   "frmLevel10.frx":C8118
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   39
      Left            =   9720
      Picture         =   "frmLevel10.frx":C966C
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   38
      Left            =   7800
      Picture         =   "frmLevel10.frx":CABC0
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   37
      Left            =   7800
      Picture         =   "frmLevel10.frx":CC114
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   36
      Left            =   9720
      Picture         =   "frmLevel10.frx":CD668
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   35
      Left            =   9720
      Picture         =   "frmLevel10.frx":CEBBC
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   34
      Left            =   9720
      Picture         =   "frmLevel10.frx":D0110
      Stretch         =   -1  'True
      Top             =   4800
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   33
      Left            =   9720
      Picture         =   "frmLevel10.frx":D1664
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   32
      Left            =   9720
      Picture         =   "frmLevel10.frx":D2BB8
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Image imgFin 
      Height          =   735
      Left            =   6120
      Picture         =   "frmLevel10.frx":D410C
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   31
      Left            =   3960
      Picture         =   "frmLevel10.frx":D512F
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   30
      Left            =   3960
      Picture         =   "frmLevel10.frx":D6683
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   29
      Left            =   7800
      Picture         =   "frmLevel10.frx":D7BD7
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   27
      Left            =   5880
      Picture         =   "frmLevel10.frx":D912B
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   26
      Left            =   3960
      Picture         =   "frmLevel10.frx":DA67F
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   25
      Left            =   1200
      Picture         =   "frmLevel10.frx":DBBD3
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   24
      Left            =   2040
      Picture         =   "frmLevel10.frx":DD127
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   23
      Left            =   2040
      Picture         =   "frmLevel10.frx":DE67B
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   22
      Left            =   2040
      Picture         =   "frmLevel10.frx":DFBCF
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   21
      Left            =   3960
      Picture         =   "frmLevel10.frx":E1123
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   20
      Left            =   120
      Picture         =   "frmLevel10.frx":E2677
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   19
      Left            =   2040
      Picture         =   "frmLevel10.frx":E3BCB
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   18
      Left            =   3960
      Picture         =   "frmLevel10.frx":E511F
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   17
      Left            =   2040
      Picture         =   "frmLevel10.frx":E6673
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   16
      Left            =   7800
      Picture         =   "frmLevel10.frx":E7BC7
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   15
      Left            =   120
      Picture         =   "frmLevel10.frx":E911B
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   14
      Left            =   7800
      Picture         =   "frmLevel10.frx":EA66F
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   13
      Left            =   9720
      Picture         =   "frmLevel10.frx":EBBC3
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   12
      Left            =   7800
      Picture         =   "frmLevel10.frx":ED117
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   11
      Left            =   9720
      Picture         =   "frmLevel10.frx":EE66B
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   10
      Left            =   9720
      Picture         =   "frmLevel10.frx":EFBBF
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   9
      Left            =   9720
      Picture         =   "frmLevel10.frx":F1113
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   8
      Left            =   7800
      Picture         =   "frmLevel10.frx":F2667
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   7
      Left            =   7800
      Picture         =   "frmLevel10.frx":F3BBB
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   6
      Left            =   5880
      Picture         =   "frmLevel10.frx":F510F
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   5
      Left            =   5880
      Picture         =   "frmLevel10.frx":F6663
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   4
      Left            =   3960
      Picture         =   "frmLevel10.frx":F7BB7
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   3
      Left            =   3960
      Picture         =   "frmLevel10.frx":F910B
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   2
      Left            =   2040
      Picture         =   "frmLevel10.frx":FA65F
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   2040
      Picture         =   "frmLevel10.frx":FBBB3
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   0
      Left            =   120
      Picture         =   "frmLevel10.frx":FD107
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   28
      Left            =   120
      Picture         =   "frmLevel10.frx":FE65B
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2055
   End
   Begin VB.Menu mnuO 
      Caption         =   "Options"
      Begin VB.Menu mnuRes 
         Caption         =   "Restart this Level"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "frmLevel10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.imgFin.Enabled = False
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
Me.tmrDown.Enabled = False
Me.tmrUp.Enabled = False
Me.tmrDown2.Enabled = False
Me.tmrUp2.Enabled = False
Me.tmrClose.Enabled = False
Me.tmrOpen.Enabled = False
Unload Me
frmLevel10.Show
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Image2.Visible = False
End Sub
Private Sub imgFin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.imgStart.Enabled = False
Me.tmrDown.Enabled = False
Me.tmrUp.Enabled = False
Me.tmrDown2.Enabled = False
Me.tmrUp2.Enabled = False
Me.tmrClose.Enabled = False
Me.tmrOpen.Enabled = False
Me.Timer1.Enabled = False
Me.Timer2.Enabled = False
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
End Sub

Private Sub imgSec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Image1(24).Visible = False
Me.lblD.Visible = True
End Sub

Private Sub imgSec2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Image1(44).Visible = False
End Sub

Private Sub imgStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = True
Next i

Me.lblA.Enabled = True

For i = 0 To Me.lblNext.Count - 1
Me.lblNext(i).Enabled = True
Next i

Me.tmrDown2.Enabled = True
Me.tmrDown.Enabled = True
Me.tmrClose.Enabled = True
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

Private Sub lblBlock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation
Unload Me
frmLevel10.Show

End Sub

Private Sub lblBlock2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation
Unload Me
frmLevel10.Show

End Sub

Private Sub lblBlock3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation
Unload Me
frmLevel10.Show
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
If MsgBox("Good Game! :D Play again?", vbQuestion + vbYesNo, "Mistwalkers") = vbYes Then
    Unload Me
    Call Lev
    Else
End
End If
End Sub

Private Sub Lev()
frmLevel1.Show
End Sub

Private Sub mnuQuit_Click()
Unload Me
End
End Sub

Private Sub mnuRes_Click()
Unload Me
frmLevel10.Show
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

Private Sub tmrClose_Timer()
Me.Image1(42).Left = Me.Image1(42).Left + 10.3
Me.Image1(37).Left = Me.Image1(37).Left - 10.3

If Me.Image1(42).Left >= 4920 Then
Me.tmrClose.Enabled = False
Me.tmrOpen.Enabled = True
Else
End If
End Sub

Private Sub tmrDown_Timer()
Me.Image1(6).Top = Me.Image1(6).Top + 30.5

If Me.Image1(6).Top >= 2400 Then
Me.tmrDown.Enabled = False
Me.tmrUp.Enabled = True
Else
End If
End Sub

Private Sub tmrDown2_Timer()
Me.Image1(0).Top = Me.Image1(0).Top + 35.5

If Me.Image1(0).Top >= 6600 Then
Me.tmrDown2.Enabled = False
Me.tmrUp2.Enabled = True
Else
End If
End Sub

Private Sub tmrOpen_Timer()
Me.Image1(42).Left = Me.Image1(42).Left - 10.3
Me.Image1(37).Left = Me.Image1(37).Left + 10.3

If Me.Image1(42).Left <= 3960 Then
Me.tmrOpen.Enabled = False
Me.tmrClose.Enabled = True
Else
End If
End Sub

Private Sub tmrUp_Timer()
Me.Image1(6).Top = Me.Image1(6).Top - 30.5

If Me.Image1(6).Top <= 840 Then
Me.tmrUp.Enabled = False
Me.tmrDown.Enabled = True
Else
End If
End Sub

Private Sub tmrUp2_Timer()
Me.Image1(0).Top = Me.Image1(0).Top - 35.5

If Me.Image1(0).Top <= 840 Then
Me.tmrUp2.Enabled = False
Me.tmrDown2.Enabled = True
Else
End If
End Sub
