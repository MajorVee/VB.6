VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "CODEJO~1.OCX"
Begin VB.MDIForm MDIfrmInt 
   BackColor       =   &H8000000C&
   Caption         =   "Memory Game"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10620
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   10620
      TabIndex        =   0
      Top             =   0
      Width           =   10620
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   8400
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   18000
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   7080
         Top             =   480
      End
      Begin VB.Label Lbltitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Memory Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1125
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   6615
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   10560
         X2              =   15360
         Y1              =   960
         Y2              =   960
      End
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   1560
      Top             =   2760
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "MDIfrmInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmGame.Show
End Sub

Private Sub MDIForm_Load()
ThemeIN Me
End Sub

Private Sub Timer2_Timer()
Lbltitle.Tag = Val(Lbltitle.Tag) + 1
If Lbltitle.Tag = "1" Then
Lbltitle.ForeColor = &H0&
ElseIf Lbltitle.Tag = "2" Then
Lbltitle.ForeColor = &H404040
ElseIf Lbltitle.Tag = "3" Then
Lbltitle.ForeColor = &H808080
ElseIf Lbltitle.Tag = "4" Then
Lbltitle.ForeColor = &HC0C0C0
ElseIf Lbltitle.Tag = "5" Then
Lbltitle.ForeColor = &HE0E0E0
ElseIf Lbltitle.Tag = "6" Then
Lbltitle.ForeColor = &HFFFFFF
ElseIf Lbltitle.Tag = "7" Then
Lbltitle.ForeColor = &HFFFFFF
ElseIf Lbltitle.Tag = "8" Then
Lbltitle.ForeColor = &HE0E0E0
ElseIf Lbltitle.Tag = "9" Then
Lbltitle.ForeColor = &HC0C0C0
ElseIf Lbltitle.Tag = "10" Then
Lbltitle.ForeColor = &H808080
ElseIf Lbltitle.Tag = "11" Then
Lbltitle.ForeColor = &H404040
ElseIf Lbltitle.Tag = "12" Then
Lbltitle.ForeColor = &H0&
Lbltitle.Tag = "0"
End If
End Sub
