VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "CODEJO~1.OCX"
Begin VB.Form frmNewName 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Player"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNewName.frx":0000
   ScaleHeight     =   2010
   ScaleWidth      =   6120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton lblOK 
      BackColor       =   &H008080FF&
      Caption         =   "O K A Y"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1800
      MaxLength       =   12
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   495
      Left            =   285
      TabIndex        =   2
      Top             =   480
      Width           =   1365
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   2880
      Top             =   1320
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmNewName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ThemeIN Me
End Sub

Private Sub lblOK_Click()
If Me.txtName.Text = "" Then
frmVN.Show
Exit Sub
End If

If RS.State = adStateOpen Then RS.Close
RS.Open "SELECT*FROM Players", CN, adOpenStatic, adLockOptimistic
    With RS
      .AddNew
        .Fields("Name") = Me.txtName.Text
        .Update
        End With
            Unload Me
    frmPlayers.Show

End Sub
