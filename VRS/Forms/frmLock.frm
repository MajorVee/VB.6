VERSION 5.00
Begin VB.Form frmLock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log - in"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4995
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmLock.frx":0000
   ScaleHeight     =   4935
   ScaleWidth      =   4995
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2810
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CheckBox chkPass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   3600
      Width           =   255
   End
   Begin Project1.jcbutton cmdLogin 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   4320
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      buttonstyle     =   10
      font            =   "frmLock.frx":7F01
      backcolor       =   12632256
      caption         =   "Login"
      usemaskcolor    =   -1  'True
   End
   Begin Project1.jcbutton cmdClose 
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   4320
      Width           =   1095
      _extentx        =   1931
      _extenty        =   661
      buttonstyle     =   10
      font            =   "frmLock.frx":7F29
      backcolor       =   8421631
      caption         =   "Close"
      usemaskcolor    =   -1  'True
   End
   Begin VB.Image Image4 
      Height          =   540
      Left            =   1080
      Picture         =   "frmLock.frx":7F51
      Top             =   3600
      Width           =   540
   End
   Begin VB.Image Image5 
      Height          =   540
      Left            =   1080
      Picture         =   "frmLock.frx":8211
      Top             =   2760
      Width           =   540
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   5
      FillColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   3135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      FillColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   3480
      Width           =   3135
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Option Explicit

Private Sub chkPass_Click()
If chkPass.Value = 1 Then
    txtPassword.PasswordChar = ""
    txtPassword.Font = "Century Gothic"
Else
    txtPassword.Font = "Century Gothic"
    txtPassword.PasswordChar = "•"
End If
End Sub

Private Sub cmdClose_Click()
End
End Sub

Private Sub cmdLogIn_Click()
Dim pRole As String
Dim ULogs As String
Dim CName As String
Set rs = New ADODB.Recordset
rs.Open "Select * From [User] Where Username='" & txtUsername.Text & "'", conn, adOpenStatic
If rs.RecordCount <> 0 Then
    If txtPassword.Text = rs!Password Then
        MsgBox "Username and Password Successfully Log In", vbInformation, "Access Grandted!"
            ULogs = txtUsername.Text
            pRole = rs!role
            CName = rs!CompName
            Set rs2 = New ADODB.Recordset
            rs2.Open "Select * from Logs", conn, 3, 3
            With rs2
                .AddNew
                .Fields("Username") = ULogs
                .Fields("TimeIn") = Time
                .Fields("LogDate") = Date
                .Fields("CompName") = CName
                .Update
                frmMain.Label8.Caption = rs2.Fields(0)
   
            End With
            
            frmMain.Label9.Caption = rs.Fields(4)
            frmMain.Label10.Caption = rs.Fields(5)
            frmMain.tmrLoad.Enabled = False
            frmMain.Label12.Caption = Time
          
          If frmMain.Label9.Caption = "GUEST" Then
            frmMain.mnuBD.Enabled = False
            End If
          
            Unload Me
            frmMain.Show
                   
            Set rs2 = Nothing
    
          
    Else
            MsgBox "Invalid Password, Please Try again!", vbCritical, "Log - In Error"
            txtUsername.Text = ""
            txtPassword.Text = ""
             txtUsername.SetFocus
    End If
Else
        MsgBox "Invalid Login, Please Try again!", vbCritical, "Log - In Error"
        txtUsername.Text = ""
        txtPassword.Text = ""
        txtUsername.SetFocus
        rs.Close
    End If

Set rs = Nothing
End Sub

Private Sub Form_DblClick()
'  frmMain.tmrLoad.Enabled = False
'  Unload Me
End Sub

Private Sub Form_Load()
modConnect.Connected
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdLogIn_Click
Else
End If
End Sub
