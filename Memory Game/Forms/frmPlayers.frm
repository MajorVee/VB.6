VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "CODEJO~1.OCX"
Begin VB.Form frmPlayers 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Players"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmPlayers.frx":0000
   ScaleHeight     =   5775
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.ListBox lstProfNames 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2940
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   1680
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   0
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   1560
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   3120
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   4680
      Top             =   4800
      Width           =   1455
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   8760
      Top             =   3960
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblTempName 
      Caption         =   "lblTempName"
      Height          =   375
      Left            =   7440
      TabIndex        =   2
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Players:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   1080
      Top             =   600
      Width           =   3975
   End
End
Attribute VB_Name = "frmPlayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Delete " & Me.lstProfNames.Text & "?", vbQuestion + vbYesNo, "Delete Profile") = vbYes Then
CN.Execute "delete from Players WHERE Name LIKE'" & frmPlayers.lblTempName.Caption & "'"
Call Form_Load
Exit Sub
Else
End If
End Sub

Private Sub cmdNew_Click()
Unload Me
frmNewName.Show
End Sub

Private Sub Form_Load()
ThemeIN Me
Call LoadPlayers
End Sub

Private Sub LoadPlayers()
If RS.State = adStateOpen Then RS.Close
RS.Open "SELECT*FROM Players", CN, adOpenStatic, adLockOptimistic
Me.lstProfNames.Clear
If RS.RecordCount = 0 Then Exit Sub
While Not RS.EOF
Me.lstProfNames.AddItem RS("Name")
RS.MoveNext
Wend
End Sub

Private Sub cmdOkay_Click()
frmMain.lblProfileName.Caption = Me.lblTempName.Caption
frmMain.lblWelcome.Visible = True
frmMain.Shape5.Visible = True
Unload Me
frmMain.Show
End Sub

Private Sub lblOK_Click()

End Sub

Private Sub lstProfNames_Click()
Me.cmdDelete.Enabled = True
Me.cmdOkay.Enabled = True

If RS.State = adStateOpen Then RS.Close
RS.Open "SELECT*FROM Players WHERE Name LIKE '" & Me.lstProfNames.Text & "'", CN, adOpenStatic, adLockOptimistic
Me.lblTempName.Caption = RS("Name")
End Sub

Private Sub Timer1_Timer()
LoadPlayers
End Sub
