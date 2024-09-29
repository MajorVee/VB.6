VERSION 5.00
Begin VB.Form frmLevel3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Image1"
   ClientHeight    =   7125
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmLevel3.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "frmLevel3.frx":2512C
   ScaleHeight     =   7125
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDown 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9000
      Top             =   2400
   End
   Begin VB.Timer tmrUp 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9000
      Top             =   1920
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   37
      Left            =   2160
      Picture         =   "frmLevel3.frx":2858E
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   36
      Left            =   3240
      Picture         =   "frmLevel3.frx":29AE2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   35
      Left            =   2040
      Picture         =   "frmLevel3.frx":2B036
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   34
      Left            =   120
      Picture         =   "frmLevel3.frx":2C58A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   33
      Left            =   7440
      Picture         =   "frmLevel3.frx":2DADE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   32
      Left            =   2040
      Picture         =   "frmLevel3.frx":2F032
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   31
      Left            =   120
      Picture         =   "frmLevel3.frx":30586
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   0
      Left            =   5880
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblNext 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      Height          =   735
      Index           =   1
      Left            =   5280
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image imgPass 
      Height          =   855
      Left            =   5280
      Top             =   1320
      Width           =   975
   End
   Begin VB.Image imgFin 
      Height          =   735
      Left            =   5400
      Picture         =   "frmLevel3.frx":31ADA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   30
      Left            =   120
      Picture         =   "frmLevel3.frx":32AFD
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   29
      Left            =   7560
      Picture         =   "frmLevel3.frx":34051
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   28
      Left            =   4200
      Picture         =   "frmLevel3.frx":355A5
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   27
      Left            =   2160
      Picture         =   "frmLevel3.frx":36AF9
      Stretch         =   -1  'True
      Top             =   720
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   26
      Left            =   6360
      Picture         =   "frmLevel3.frx":3804D
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   25
      Left            =   6960
      Picture         =   "frmLevel3.frx":395A1
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   24
      Left            =   7680
      Picture         =   "frmLevel3.frx":3AAF5
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   23
      Left            =   3120
      Picture         =   "frmLevel3.frx":3C049
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   22
      Left            =   5040
      Picture         =   "frmLevel3.frx":3D59D
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   21
      Left            =   120
      Picture         =   "frmLevel3.frx":3EAF1
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   20
      Left            =   120
      Picture         =   "frmLevel3.frx":40045
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   19
      Left            =   2160
      Picture         =   "frmLevel3.frx":41599
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   18
      Left            =   120
      Picture         =   "frmLevel3.frx":42AED
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   17
      Left            =   120
      Picture         =   "frmLevel3.frx":44041
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   16
      Left            =   120
      Picture         =   "frmLevel3.frx":45595
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   15
      Left            =   1560
      Picture         =   "frmLevel3.frx":46AE9
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   14
      Left            =   120
      Picture         =   "frmLevel3.frx":4803D
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   13
      Left            =   7560
      Picture         =   "frmLevel3.frx":49591
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   12
      Left            =   2760
      Picture         =   "frmLevel3.frx":4AAE5
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   11
      Left            =   3600
      Picture         =   "frmLevel3.frx":4C039
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   10
      Left            =   7560
      Picture         =   "frmLevel3.frx":4D58D
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   9
      Left            =   120
      Picture         =   "frmLevel3.frx":4EAE1
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   8
      Left            =   120
      Picture         =   "frmLevel3.frx":50035
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   7
      Left            =   7560
      Picture         =   "frmLevel3.frx":51589
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   6
      Left            =   7560
      Picture         =   "frmLevel3.frx":52ADD
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   5
      Left            =   6840
      Picture         =   "frmLevel3.frx":54031
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   4
      Left            =   3600
      Picture         =   "frmLevel3.frx":55585
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   3
      Left            =   5640
      Picture         =   "frmLevel3.frx":56AD9
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   2
      Left            =   3720
      Picture         =   "frmLevel3.frx":5802D
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   1
      Left            =   4920
      Picture         =   "frmLevel3.frx":59581
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   615
      Index           =   0
      Left            =   6960
      Picture         =   "frmLevel3.frx":5AAD5
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Image imgStart 
      Height          =   615
      Left            =   120
      Picture         =   "frmLevel3.frx":5C029
      Stretch         =   -1  'True
      Top             =   5880
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
Attribute VB_Name = "frmLevel3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.imgFin.Enabled = False
Me.imgPass.Enabled = False
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = False
Next i
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.tmrDown.Enabled = False
Me.tmrUp.Enabled = False
MsgBox "YOU'RE DEAD! TRY AGAIN!", vbExclamation, "Mistwalkers"
Unload Me
frmLevel3.Show
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
Me.imgPass.Enabled = True
For i = 0 To Me.Image1.Count - 1
Me.Image1(i).Enabled = True
Next i

Me.tmrDown.Enabled = True
End Sub

Private Sub lblNext_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
frmLevel4.Show
End Sub

Private Sub mnuQuit_Click()
If MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Mistwalkers") = vbYes Then
    End
End If
End Sub

Private Sub mnuRes_Click()
Unload Me
frmLevel3.Show
End Sub

Private Sub tmrDown_Timer()
Me.Image1(1).Top = Me.Image1(1).Top + 10.1

If Me.Image1(1).Top >= 3480 Then
Me.tmrDown.Enabled = False
Me.tmrUp.Enabled = True
Else
End If
End Sub

Private Sub tmrUp_Timer()
Me.Image1(1).Top = Me.Image1(1).Top - 10.1
If Me.Image1(1).Top <= 2280 Then
Me.tmrUp.Enabled = False
Me.tmrDown.Enabled = True
Else
End If
End Sub
