VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "CODEJO~1.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Game"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   7410
   ScaleWidth      =   6165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDown 
      Interval        =   1
      Left            =   6240
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   6240
      Top             =   4920
   End
   Begin VB.CommandButton cmdEnd 
      BackColor       =   &H0080FF80&
      Caption         =   "E x i t"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   6240
      Width           =   2415
   End
   Begin VB.CommandButton cmdHS 
      BackColor       =   &H0080C0FF&
      Caption         =   "Score Board"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   4
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFF80&
      Caption         =   "Select Profile"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   3
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H008080FF&
      Caption         =   "P L A Y"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Memory Game"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   525
      TabIndex        =   6
      Top             =   -840
      Width           =   5025
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   1680
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   1680
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   1680
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label lblWelcome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   6240
      Top             =   5640
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblProfileName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   1680
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
frmPlayers.Show
Unload Me
End Sub

Private Sub cmdEnd_Click()
If MsgBox("Do you want to exit?", vbQuestion + vbYesNo, "Memory Game") = vbYes Then
    End
Else
End If
End Sub

Private Sub cmdHS_Click()
Unload Me
frmTopScore.Show
End Sub

Private Sub cmdStart_Click()
If Me.lblProfileName.Caption = "" Then
frmChoose.Show
Exit Sub
End If
frmGame.Label5.Caption = Me.lblProfileName.Caption
Unload Me
frmGame.Show
End Sub

Private Sub Form_Load()
ThemeIN Me
Me.lblProfileName.Caption = ""
End Sub

Private Sub Timer1_Timer()
Label1.Tag = Val(Label1.Tag) + 1
If Label1.Tag = "1" Then
Label1.ForeColor = &H0&
ElseIf Label1.Tag = "2" Then
Label1.ForeColor = &H404040
ElseIf Label1.Tag = "3" Then
Label1.ForeColor = &H808080
ElseIf Label1.Tag = "4" Then
Label1.ForeColor = &H8080FF
ElseIf Label1.Tag = "5" Then
Label1.ForeColor = &HE0E0E0
ElseIf Label1.Tag = "6" Then
Label1.ForeColor = &HC0C0FF
ElseIf Label1.Tag = "7" Then
Label1.ForeColor = &HFFFFFF
ElseIf Label1.Tag = "8" Then
Label1.ForeColor = &HC0FFC0
ElseIf Label1.Tag = "9" Then
Label1.ForeColor = &H80FF80
ElseIf Label1.Tag = "10" Then
Label1.ForeColor = &HFF00&
ElseIf Label1.Tag = "11" Then
Label1.ForeColor = &HC000&
ElseIf Label1.Tag = "12" Then
Label1.ForeColor = &H8000&
ElseIf Label1.Tag = "13" Then
Label1.ForeColor = &H0&
Label1.Tag = "0"
End If
End Sub

Private Sub tmrDown_Timer()
Me.Label1.Top = Me.Label1.Top + 10.1

If Me.Label1.Top >= 1320 Then
Me.tmrDown.Enabled = False
Me.Timer1.Enabled = True
Else
End If
End Sub
