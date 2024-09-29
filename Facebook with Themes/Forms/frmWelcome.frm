VERSION 5.00
Begin VB.Form frmWelcome 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrLoad 
      Interval        =   2000
      Left            =   2640
      Top             =   2760
   End
   Begin VB.Label lblIndacator 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   2535
   End
   Begin VB.Label lblMsgBoxCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sorry, either your email or password is invalid."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   435
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   6810
   End
   Begin VB.Image imgButton 
      Height          =   285
      Left            =   4080
      Picture         =   "frmWelcome.frx":0000
      Top             =   1800
      Width           =   1065
   End
   Begin VB.Label lblbWelcome 
      Caption         =   "Label1"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2760
      Width           =   615
   End
End
Attribute VB_Name = "frmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmLogin.Enabled = False
Me.lblIndacator.Visible = False
Me.lblbWelcome.Visible = False
Me.lblbWelcome.Caption = "1"

'=======================Transparent Form=================
MakeTransparent Me.hwnd, 210

'=======================Pass=======================
Me.lblbWelcome.Caption = "5"
End Sub

Private Sub imgButton_Click()
frmLogin.Enabled = True
Unload Me
frmLogin.Show
End Sub

