VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "CODEJO~1.OCX"
Begin VB.MDIForm MDIfrmFB 
   BackColor       =   &H8000000F&
   Caption         =   "F A C E B O O K"
   ClientHeight    =   8385
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14280
   LinkTopic       =   "MDIForm1"
   Moveable        =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1335
      ScaleWidth      =   20250
      TabIndex        =   0
      Top             =   0
      Width           =   20250
      Begin VB.Frame Frame3 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   15000
         TabIndex        =   4
         Top             =   360
         Width           =   2895
         Begin VB.CommandButton Command1 
            Caption         =   "Jean"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Home"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   1560
            TabIndex        =   5
            Top             =   120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Line Line3 
            BorderColor     =   &H80000000&
            Visible         =   0   'False
            X1              =   1440
            X2              =   1440
            Y1              =   0
            Y2              =   720
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   5760
         Top             =   720
      End
      Begin VB.ListBox List1 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1095
         Left            =   18000
         TabIndex        =   2
         Top             =   120
         Width           =   2175
      End
      Begin VB.Image Image2 
         Height          =   975
         Left            =   13800
         Picture         =   "MDIfrmFB.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jean L. Reyes"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   660
         Left            =   9240
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Line Line1 
         Visible         =   0   'False
         X1              =   8880
         X2              =   13680
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image5 
         Height          =   705
         Left            =   12960
         Picture         =   "MDIfrmFB.frx":133A5
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Image Image1 
         Height          =   1320
         Left            =   0
         Picture         =   "MDIfrmFB.frx":16B1C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1425
      End
      Begin VB.Label Lbltitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "acebook"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   72
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1680
         Left            =   1440
         TabIndex        =   1
         Top             =   -120
         Width           =   6390
      End
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   600
      Top             =   1440
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuFB 
      Caption         =   "FACEBOOK"
      Begin VB.Menu mnuHome 
         Caption         =   "P R O F I L E"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuLO 
         Caption         =   "CLOSE FACEBOOK"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "MDIfrmFB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
frmProfile.Show
End Sub

Private Sub mnuHome_Click()
frmThemes.Show
Label1.Visible = True
Image5.Visible = True
Image2.Visible = True
Line1.Visible = True
Line3.Visible = True
Command1.Visible = True
Command2.Visible = True
End Sub
Private Sub MDIForm_Load()
LoadTheme
End Sub
Private Sub LoadTheme()
Me.List1.AddItem "BASIC"
Me.List1.AddItem "Blink"
Me.List1.AddItem "Boost"
Me.List1.AddItem "Blue Sea"
Me.List1.AddItem "BumbleBee"
Me.List1.AddItem "Cosmo"
Me.List1.AddItem "Cocoy"
Me.List1.AddItem "Fresco"
Me.List1.AddItem "Fusion VS"
Me.List1.AddItem "Green Grass"
Me.List1.AddItem "Harvest"
Me.List1.AddItem "Hex"
Me.List1.AddItem "Mac OS-X"
Me.List1.AddItem "Manzanas"
Me.List1.AddItem "PinkLoop"
Me.List1.AddItem "Red Dragon"
Me.List1.AddItem "Rogue"
Me.List1.AddItem "Trippin"
Me.List1.AddItem "Vincent"
Me.List1.AddItem "VS7"
End Sub

Private Sub List1_Click()
ThemeIN Me
End Sub

Private Sub mnuClose_Click()
End
End Sub

Private Sub mnuLO_Click()
End
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
