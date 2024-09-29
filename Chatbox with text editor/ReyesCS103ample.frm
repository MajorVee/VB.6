VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000008&
   Caption         =   "Form1"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8940
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "ReyesCS103ample.frx":0000
   ScaleHeight     =   4755
   ScaleMode       =   0  'User
   ScaleWidth      =   8940
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   435
      ItemData        =   "ReyesCS103ample.frx":1A70C
      Left            =   6960
      List            =   "ReyesCS103ample.frx":1A737
      TabIndex        =   18
      Text            =   "Theme"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFC0FF&
      Height          =   2265
      ItemData        =   "ReyesCS103ample.frx":1A7B3
      Left            =   4200
      List            =   "ReyesCS103ample.frx":1A7BA
      TabIndex        =   17
      Top             =   2160
      Width           =   4455
   End
   Begin VB.CheckBox chkItalic 
      BackColor       =   &H0080FFFF&
      Caption         =   "&Italic"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   16
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "&Display"
      Default         =   -1  'True
      Height          =   735
      Left            =   4920
      TabIndex        =   15
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   735
      Left            =   6960
      TabIndex        =   14
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "&Clear"
      Height          =   735
      Left            =   4920
      TabIndex        =   13
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Style"
      Height          =   2415
      Left            =   1800
      TabIndex        =   10
      Top             =   2160
      Width           =   2175
      Begin VB.CheckBox chkUnderline 
         BackColor       =   &H0080FFFF&
         Caption         =   "Underline"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CheckBox chkBold 
         BackColor       =   &H0080FFFF&
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Color:"
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
      Begin VB.OptionButton optRed 
         BackColor       =   &H00C0C000&
         Caption         =   "&Red"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton optBlack 
         BackColor       =   &H00C0C000&
         Caption         =   "&Black"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1335
      End
      Begin VB.OptionButton optBlue 
         BackColor       =   &H00C0C000&
         Caption         =   "&Blue"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.OptionButton optGreen 
         BackColor       =   &H00C0C000&
         Caption         =   "&Green"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "USER INPUT"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtMessage 
         BackColor       =   &H0080FFFF&
         Height          =   555
         Left            =   1680
         TabIndex        =   4
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H0080FFFF&
         Height          =   435
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message:"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:"
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1320
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Date:November 16, 2015
'Programmer:Arvie Reyes
'Description:This program uses control to change the properties of text


Private Sub chkBold_Click()
'Change the message text to/from bold

Me.List1.Font.Bold = chkBold.Value
End Sub

Private Sub chkItalic_Click()
'Change the message text to/from italic

Me.List1.Font.Italic = Me.chkItalic.Value
End Sub

Private Sub chkUnderline_Click()
'Change the message text to/from underline

Me.List1.Font.Underline = Me.chkUnderline.Value
End Sub

Private Sub cmdClear_Click()
Me.txtName.Text = ""
Me.txtMessage.Text = ""
End Sub

Private Sub cmdDisplay_Click()
'Display error message if textboxes aren't filled
If Me.txtName.Text = "" And Me.txtMessage.Text = "" Then
MsgBox "You failed to input your name and message", vbExclamation
Exit Sub
End If
If Me.txtMessage.Text = "" Then
MsgBox "You have failed to input your Message", vbExclamation
Exit Sub
End If
If Me.txtName.Text = "" Then
MsgBox "You have failed to input your Name", vbExclamation
Exit Sub
End If

'Display the text in the message area

Me.List1.AddItem Me.txtName.Text & ":" & Me.txtMessage.Text
cmdClear_Click
Me.txtName.SetFocus
End Sub

Private Sub cmdExit_Click()
'Exit the project

End
End Sub
Private Sub Combo1_Click()
If Me.Combo1.Text = "Winter" Then
Me.txtName.BackColor = &H80FF80
Me.txtMessage.BackColor = &HFFC0FF
Me.Picture = LoadPicture(App.Path & "/Design/Winter2.jpg ")

ElseIf Me.Combo1.Text = "Spring" Then
Me.txtName.BackColor = &HFF80FF
Me.txtMessage.BackColor = &HC0FFC0
Me.Picture = LoadPicture(App.Path & "/Design/Spring2.jpg")

ElseIf Me.Combo1.Text = "Moon" Then
Me.txtName.BackColor = &HC0E0FF
Me.txtMessage.BackColor = &HFFFFC0
Me.Picture = LoadPicture(App.Path & "/Design/Night2.jpg")

ElseIf Me.Combo1.Text = "Beach" Then
Me.Picture = LoadPicture(App.Path & "/Design/4.jpg")

ElseIf Me.Combo1.Text = "Eiffel Tower" Then
Me.Picture = LoadPicture(App.Path & "/Design/Paris.jpg")

ElseIf Me.Combo1.Text = "Painting 1" Then
Me.Picture = LoadPicture(App.Path & "/Design/1.jpg")

ElseIf Me.Combo1.Text = "Painting 2" Then
Me.Picture = LoadPicture(App.Path & "/Design/3.jpg")

ElseIf Me.Combo1.Text = "Painting 3" Then
Me.Picture = LoadPicture(App.Path & "/Design/5.jpg")

ElseIf Me.Combo1.Text = "All Star" Then
Me.Picture = LoadPicture(App.Path & "/Design/Dota.jpg")

ElseIf Me.Combo1.Text = "Dota 2" Then
Me.Picture = LoadPicture(App.Path & "/Design/Dota2.jpg")

ElseIf Me.Combo1.Text = "Fence" Then
Me.Picture = LoadPicture(App.Path & "/Design/Fence.jpg")

ElseIf Me.Combo1.Text = "Marian" Then
Me.Picture = LoadPicture(App.Path & "/Design/Marian.gif")

ElseIf Me.Combo1.Text = "Nature" Then
Me.Picture = LoadPicture(App.Path & "/Design/Nature.jpg")

End If
End Sub

Private Sub Form_Load()

End Sub

Private Sub optBlack_Click()
'Make message text black

Me.List1.ForeColor = vbBlack
End Sub

Private Sub optBlue_Click()
'Make message text blue

Me.List1.ForeColor = vbBlue
End Sub

Private Sub optGreen_Click()
'Make message text green

Me.List1.ForeColor = vbGreen
End Sub

Private Sub optRed_Click()
'Make message text red

Me.List1.ForeColor = vbRed
End Sub
