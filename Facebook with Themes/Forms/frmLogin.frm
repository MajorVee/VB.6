VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00800000&
   Caption         =   "Facebook Login"
   ClientHeight    =   2925
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   12720
      Top             =   3360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13320
      Top             =   4320
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7920
      TabIndex        =   1
      Top             =   3960
      Width           =   5175
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      TabIndex        =   0
      Top             =   4560
      Width           =   5175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   420
      Left            =   6000
      TabIndex        =   6
      Top             =   4560
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   420
      Left            =   6480
      TabIndex        =   5
      Top             =   3960
      Width           =   1050
   End
   Begin VB.Label Lbltitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Facebook Login"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   7680
      TabIndex        =   4
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   3
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Menu mnuF 
      Caption         =   "Facebook"
      Begin VB.Menu mnuCF 
         Caption         =   "Close Facebook"
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub Enter()
Me.Label2.Caption = Me.txtUser.Text & Me.txtPass.Text
Me.Timer1.Enabled = True
End Sub

Private Sub Command1_Click()
Me.Label2.Caption = Me.txtUser.Text & Me.txtPass.Text
Me.Timer1.Enabled = True
End Sub

Private Sub Form_Load()
Me.Timer1.Enabled = False
Me.Label1.Visible = False
Me.Label2.Visible = False
Me.txtPass.PasswordChar = "*"
Me.txtPass.FontSize = "20"
Me.Label1.Caption = ""
Me.Label2.Caption = ""
If RS.State = adStateOpen Then RS.Close
RS.Open " SELECT*From AdminT", CN, aropenstatic, adLockOptimistic

Me.Label1.Caption = RS("User") & RS("Password")

End Sub

Private Sub mnuCF_Click()
End
End Sub

Private Sub Timer1_Timer()
If Me.Label2.Caption = Me.Label1.Caption Then
MsgBox "Welcome Jean!", vbOKOnly + vbInformation, "Facebook"

MDIfrmFB.Show
Unload Me
Exit Sub
Else
Me.txtUser.Text = ""
Me.txtPass.Text = ""
Me.Label2.Caption = ""
MsgBox "Access Denied!", vbExclamation
frmWelcome.Show
Me.Timer1.Enabled = False
End If
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

Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call Enter
Else
End If

End Sub

'Private Sub txtUser_KeyPress(KeyAscii As Integer)
'Select Case KeyAscii
    'Case 97 To 122
      'KeyAscii = KeyAscii - 32
  'End Select
'End Sub

Private Sub txtUser_LostFocus()
Me.txtUser.Text = UCase(Me.txtUser.Text)
End Sub

