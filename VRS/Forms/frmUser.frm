VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUser 
   BackColor       =   &H80000007&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User Accounts"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmUser.frx":0000
   ScaleHeight     =   6570
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1440
      Top             =   5160
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "User Account Information"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6495
      Begin VB.TextBox txtConfirm 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "•"
         TabIndex        =   6
         Top             =   1680
         Width           =   4455
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "•"
         TabIndex        =   5
         Top             =   1080
         Width           =   4455
      End
      Begin VB.TextBox txtUsername 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1560
         TabIndex        =   4
         ToolTipText     =   "Ex. John2014"
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1560
         TabIndex        =   3
         Top             =   2280
         Width           =   4455
      End
      Begin VB.ComboBox cboRole 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         ItemData        =   "frmUser.frx":5601
         Left            =   1560
         List            =   "frmUser.frx":5603
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2880
         Width           =   4455
      End
      Begin VB.CheckBox chkPass 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   6120
         TabIndex        =   1
         Top             =   1440
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtDate 
         Height          =   405
         Left            =   1560
         TabIndex        =   7
         Top             =   3480
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   14737632
         Format          =   105512960
         CurrentDate     =   42979
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   1140
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   1200
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Role:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   480
         TabIndex        =   9
         Top             =   2880
         Width           =   555
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Reg:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   3480
         Width           =   1185
      End
   End
   Begin MSComctlLib.ListView lvUserAccount 
      Height          =   4095
      Left            =   6720
      TabIndex        =   14
      Top             =   1200
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12632256
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User ID"
         Object.Width           =   18
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Username"
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   4058
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Role"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date Registered"
         Object.Width           =   3951
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Password"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Confirm"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":5605
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   149
      ImageHeight     =   57
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":5B9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":9CB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":D5F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1154F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":14F37
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1947B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":1CF19
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":20E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":247AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":2836A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "User Accounts List"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1800
      TabIndex        =   22
      Top             =   360
      Width           =   3195
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Counter"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8640
      TabIndex        =   20
      Top             =   480
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   600
      Picture         =   "frmUser.frx":2BAEE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   7440
      TabIndex        =   19
      Top             =   5760
      Width           =   1050
   End
   Begin VB.Image imgUpdate 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   6480
      Picture         =   "frmUser.frx":2DE1C
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   2040
      TabIndex        =   18
      Top             =   5760
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   4920
      TabIndex        =   17
      Top             =   5760
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   10080
      TabIndex        =   16
      Top             =   5760
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   12720
      TabIndex        =   15
      Top             =   5760
      Width           =   1005
   End
   Begin VB.Image imgAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   1200
      Picture         =   "frmUser.frx":31590
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Image imgSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   3840
      Picture         =   "frmUser.frx":34EC2
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Image imgDel 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   9120
      Picture         =   "frmUser.frx":3889A
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Image imgRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   11760
      Picture         =   "frmUser.frx":3C328
      Top             =   5520
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   0
      TabIndex        =   21
      Top             =   240
      Width           =   20655
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim usr As String
Dim role As String

Private Sub chkPass_Click()
If chkPass.Value = 1 Then
    txtPassword.PasswordChar = ""
    txtPassword.Font = "Arial"
    txtConfirm.PasswordChar = ""
    txtConfirm.Font = "Arial"
Else
    txtPassword.Font = "Arial"
    txtPassword.PasswordChar = "•"
    txtConfirm.Font = "Arial"
    txtConfirm.PasswordChar = "•"
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()

modConnect.Connected
Call RefAccount
Call PopRole
imgUpdate.Enabled = False
Call UserLock
End Sub

Public Sub UserLock()
frmUser.txtUsername.Enabled = False
frmUser.txtPassword.Enabled = False
frmUser.txtConfirm.Enabled = False
frmUser.txtName.Enabled = False
frmUser.cboRole.Enabled = False
frmUser.dtDate.Enabled = False
frmUser.chkPass.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
End Sub
Private Sub imgAdd_Click()
        Call imgRefresh_Click
        imgDel.Enabled = False
        imgUpdate.Enabled = False
        Call UserMan_Clear
frmUser.txtUsername.Enabled = True
frmUser.txtPassword.Enabled = True
frmUser.txtConfirm.Enabled = True
frmUser.txtName.Enabled = True
frmUser.cboRole.Enabled = True
frmUser.dtDate.Enabled = True
frmUser.chkPass.Enabled = True
End Sub

Private Sub imgAdd_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(1).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
End Sub

Private Sub imgDel_Click()
If usr = vbNullString Then
            MsgBox "Please choose a record in the list.", vbExclamation, "Warning!"
        Else
        
        If MsgBox("Delete?", vbQuestion + vbYesNo) = vbYes Then
            
            Set rs = New ADODB.Recordset
            rs.Open "Select * from [User] where Username='" & usr & "'", conn, adOpenKeyset, adLockPessimistic
            With rs
                .Delete
                Call RefAccount
'                Call NoUser
                .Close
            End With
            Set rs = Nothing
            MsgBox "Record Successfully Deleted!", vbInformation, "Success Deleted!"
        End If
End If

    Call UserMan_Clear
    Call RefAccount
    Call UserLock

End Sub

Private Sub imgDel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(5).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
End Sub

Private Sub imgRefresh_Click()
        imgSave.Enabled = True
        Call UserMan_Clear
        Call RefAccount
        Call UserLock
End Sub

Private Sub imgRefresh_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(7).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
End Sub

Private Sub imgSave_Click()
 If txtUsername.Text = "" Or txtPassword.Text = "" Or txtConfirm.Text = "" Or txtName.Text = "" Or cboRole.Text = "" Or dtDate.Value = 0 Then
            MsgBox "Some of your fields is empty. Please complete the information.", vbExclamation + vbOKOnly, "Warning"
'            txtUsername.SetFocus
        ElseIf txtPassword.Text <> txtConfirm.Text Then
            MsgBox "Your password did not match, Please re-enter.", vbExclamation + vbOKOnly, "Warning"
            txtPassword.Text = ""
            txtConfirm.Text = ""
            txtPassword.SetFocus
        Else
            Set rs = New ADODB.Recordset
            rs.Open "Select * from [User]", conn, adOpenKeyset, adLockPessimistic
                With rs
                    Dim a As String
                    rs.MoveLast
                    a = rs.Fields("UserID")
                    a = a + 1
                    .AddNew
                    .Fields("UserID") = "00" & a
                    .Fields("Username") = txtUsername.Text
                    .Fields("Password") = txtPassword.Text
                    .Fields("Confirm") = txtConfirm.Text
                    .Fields("CompName") = txtName.Text
'                    .Fields("Email") = txtEmail.Text
                    .Fields("Role") = cboRole.Text
                    .Fields("DateReg") = dtDate.Value
                    .Update
                    Call RefAccount
'                    Call NoUser
                     End With
            'rs.Close
            Set rs = Nothing
            MsgBox "Record Successfully Saved!", vbInformation, "Success Saved!"
            Call UserMan_Clear
            
        End If
Call RefAccount
Call UserLock
End Sub

Private Sub imgSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(3).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
End Sub

Private Sub imgUpdate_Click()
  If txtUsername.Text = "" Or txtPassword.Text = "" Or txtConfirm.Text = "" Or txtName.Text = "" Or cboRole.Text = "" Or dtDate.Value = 0 Then
            MsgBox "Some of your fields is empty. Please complete the information.", vbExclamation + vbOKOnly, "Warning"
'            txtUsername.SetFocus
        Else
            Set rs = New ADODB.Recordset
                rs.Open "Select * from [User] where Username='" & usr & "'", conn, adOpenDynamic, adLockOptimistic
            With rs
                !UserName = txtUsername.Text
                !Password = txtPassword.Text
                !Confirm = txtConfirm.Text
                !CompName = txtName.Text
                !role = cboRole.Text
                !DateReg = dtDate.Value
                rs.Update
            Call RefAccount
'            Call NoUser
            End With
            Set rs = Nothing
            MsgBox "Record Successfully Updated!", vbInformation, "Success Updated!"
    End If
Call UserMan_Clear
Call UserLock
End Sub

Private Sub imgUpdate_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(9).Picture
End Sub

Public Sub PopRole()
Set rs = New ADODB.Recordset
rs.Open "Select RoleName from Role Order by RoleName ASC", conn, 3, 3
Do While Not rs.EOF
    frmUser.cboRole.AddItem rs!RoleName
rs.MoveNext
Loop
End Sub

Private Sub lvUserAccount_Click()
frmUser.txtUsername.Enabled = True
frmUser.txtPassword.Enabled = True
frmUser.txtConfirm.Enabled = True
frmUser.txtName.Enabled = True
frmUser.cboRole.Enabled = True
frmUser.dtDate.Enabled = True
frmUser.chkPass.Enabled = True
imgUpdate.Enabled = True
imgSave.Enabled = False
imgDel.Enabled = True

    On Error Resume Next
        usr = lvUserAccount.SelectedItem.SubItems(1)
        txtUsername.Text = lvUserAccount.SelectedItem.SubItems(1)
        txtName.Text = lvUserAccount.SelectedItem.SubItems(2)
        cboRole.Text = lvUserAccount.SelectedItem.SubItems(3)
        dtDate.Value = lvUserAccount.SelectedItem.SubItems(4)
        txtPassword.Text = lvUserAccount.SelectedItem.SubItems(5)
        txtConfirm.Text = lvUserAccount.SelectedItem.SubItems(6)
End Sub

Private Sub lvUserAccount_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
End Sub

Private Sub Timer1_Timer()
If val(Me.lvUserAccount.ListItems.Count) > 1 Then
Me.Label15.Caption = "There are " & Me.lvUserAccount.ListItems.Count & " items found in the list."
ElseIf Me.lvUserAccount.ListItems.Count = 1 Then
Me.Label15.Caption = "There is " & Me.lvUserAccount.ListItems.Count & " item found in the list."
Else
Me.Label15.Caption = "No item found in the list."
End If
End Sub
