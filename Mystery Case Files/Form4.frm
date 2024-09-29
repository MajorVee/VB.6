VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000007&
   Caption         =   "Form4"
   ClientHeight    =   10440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20055
   LinkTopic       =   "Form4"
   ScaleHeight     =   10440
   ScaleWidth      =   20055
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H008080FF&
      Caption         =   "Restart"
      Height          =   375
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9840
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H8000000A&
      Caption         =   "<<Back to Menu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15960
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9600
      UseMaskColor    =   -1  'True
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Things to Look! Hurry!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3375
      Left            =   16200
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Image Image4 
         Height          =   2340
         Left            =   1680
         Picture         =   "Form4.frx":0000
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1980
      End
      Begin VB.Label lblCap 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cap"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   8
         Top             =   2520
         Width           =   570
      End
      Begin VB.Label lblTrum 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trumpet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   480
         TabIndex        =   7
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label lblSock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   1560
         Width           =   690
      End
      Begin VB.Label lblSkate 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label lblSax 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saxophone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   1650
      End
   End
   Begin VB.Image imgCap 
      Height          =   420
      Left            =   5160
      Picture         =   "Form4.frx":82BC0
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   405
   End
   Begin VB.Image imgTrumpet 
      Height          =   1020
      Left            =   11160
      Picture         =   "Form4.frx":FE8C4
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1020
   End
   Begin VB.Image imgSock 
      Height          =   600
      Left            =   6120
      Picture         =   "Form4.frx":13826F
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   600
   End
   Begin VB.Image imgSkate 
      Height          =   420
      Left            =   480
      Picture         =   "Form4.frx":13DC6E
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   540
   End
   Begin VB.Image imgSax 
      Height          =   735
      Left            =   8880
      Picture         =   "Form4.frx":15AA27
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   17880
      TabIndex        =   3
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   375
      Left            =   16560
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   2880
      Left            =   16560
      Picture         =   "Form4.frx":1A1BB5
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   3045
   End
   Begin VB.Image imgEnd 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   19200
      Picture         =   "Form4.frx":1B5148
      Stretch         =   -1  'True
      Top             =   9600
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   3675
      Left            =   15840
      Picture         =   "Form4.frx":1B9A63
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   4500
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   11520
      Left            =   -2040
      Picture         =   "Form4.frx":1CA3DD
      Top             =   0
      Width           =   20400
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuPlay 
         Caption         =   "How to Play the Game"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
Form1.Show
End Sub

Private Sub cmdStart_Click()
Unload Me
Form4.Show
End Sub

Private Sub imgCap_Click()
Me.imgCap.Visible = False
Me.lblCap.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgEnd_Click()
'To end the Program
If MsgBox("Are you sure you want to stop playing?", vbQuestion + vbYesNo, "System") = vbYes Then
    End
End If
End Sub

Private Sub imgSax_Click()
Me.imgSax.Visible = False
Me.lblSax.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgSkate_Click()
Me.imgSkate.Visible = False
Me.lblSkate.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgSock_Click()
Me.imgSock.Visible = False
Me.lblSock.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgTrumpet_Click()
Me.imgTrumpet.Visible = False
Me.lblTrum.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub lblScore_Change()
If Val(Me.lblScore.Caption) = 5 Then
 MsgBox "Ikaw na! Edi wow! Ikaw na magaling _=="
Exit Sub
End If
End Sub

Private Sub mnuPlay_Click()
Unload Me
Form6.Show
End Sub
