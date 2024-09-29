VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00800000&
   Caption         =   "Facebook Login"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   7440
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6600
      Top             =   1920
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
      Left            =   2040
      TabIndex        =   1
      Top             =   2400
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
      Left            =   2040
      TabIndex        =   0
      Top             =   3000
      Width           =   5175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   2880
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   2400
      Width           =   915
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
      Left            =   1680
      TabIndex        =   4
      Top             =   480
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
      Left            =   2160
      TabIndex        =   3
      Top             =   1440
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   4215
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

Private Sub Timer1_Timer()
If Me.Label2.Caption = Me.Label1.Caption Then
MsgBox "Successfully Login! Welcome to Facebook Jean L. Reyes!", vbOKOnly + vbInformation, "Facebook"

MDIfrmFB.Show
Unload Me
Exit Sub
Else
Me.txtUser.Text = ""
Me.txtPass.Text = ""
Me.Label2.Caption = ""
MsgBox "Access Denied!", vbExclamation
Me.txtUser.SetFocus
Me.Timer1.Enabled = False
End If
End Sub

Private Sub mnuHome_Click()
frmThemes.Show
End Sub
Private Sub MDIForm_Load()
LoadTheme
End Sub
Private Sub LoadTheme()
Me.List1.AddItem "Anti Xero"
Me.List1.AddItem "BASIC"
Me.List1.AddItem "Blink"
Me.List1.AddItem "Boost"
Me.List1.AddItem "Blue Sea"
Me.List1.AddItem "BumbleBee"
Me.List1.AddItem "Cosmo"
Me.List1.AddItem "Cocoy"
Me.List1.AddItem "Dark Revo"
Me.List1.AddItem "Felicity"
Me.List1.AddItem "Fresco"
Me.List1.AddItem "Fusion VS"
Me.List1.AddItem "Green Grass"
Me.List1.AddItem "Harvest"
Me.List1.AddItem "Hex"
Me.List1.AddItem "HomeRedo"
Me.List1.AddItem "Hover"
Me.List1.AddItem "Mac OS-X"
Me.List1.AddItem "Manzanas"
Me.List1.AddItem "Native Grey"
Me.List1.AddItem "PinkLoop"
Me.List1.AddItem "Red Dragon"
Me.List1.AddItem "Rogue"
Me.List1.AddItem "Trippin"
Me.List1.AddItem "Vincent"
Me.List1.AddItem "VS7"
Me.List1.AddItem "Windows XP"
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

