VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "CODEJO~1.OCX"
Begin VB.Form frmTopScore 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Top Players"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmTopScore.frx":0000
   ScaleHeight     =   6285
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   1
      Top             =   5640
      Width           =   1335
   End
   Begin MSComctlLib.ListView lstViewScore 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7858
      SortKey         =   1
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   12632256
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NAME"
         Object.Width           =   6950
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "SCORE"
         Object.Width           =   5539
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scoreboard"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   2490
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   5280
      Top             =   4560
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   1920
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmTopScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'Cancel Button
Unload Me
frmMain.Show
End Sub
Private Sub Form_Load()
ThemeIN Me
Call LoadAllScore
End Sub
Private Sub LoadAllScore()
On Error Resume Next
If RS.State = adStateOpen Then RS.Close
RS.Open "SELECT*FROM Players", CN, adOpenStatic, adLockOptimistic
Me.lstViewScore.ListItems.Clear
If RS.RecordCount = 0 Then Exit Sub
While Not RS.EOF
 If Not RS("Score") = "0" Or RS("Score") = "" Then
    Set lst = Me.lstViewScore.ListItems.Add(, , RS("Name"))
        lst.SubItems(1) = RS("Score")
    End If
        RS.MoveNext
Wend
End Sub


