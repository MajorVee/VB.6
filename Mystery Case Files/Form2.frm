VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000007&
   Caption         =   "Form2"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   14490
   LinkTopic       =   "Form2"
   ScaleHeight     =   9330
   ScaleWidth      =   14490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H008080FF&
      Caption         =   "Restart"
      Height          =   375
      Left            =   11040
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8880
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Left            =   9600
      Top             =   8280
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
      Height          =   495
      Left            =   10680
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8280
      UseMaskColor    =   -1  'True
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Caption         =   "Found Objects! Good Job!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   9600
      TabIndex        =   1
      Top             =   4920
      Width           =   4815
      Begin VB.Image Image2 
         Height          =   2745
         Left            =   1320
         Picture         =   "Form2.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3000
      End
      Begin VB.Label lblScore 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
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
         Height          =   2415
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Objects to find! Hurry!"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   4695
      Left            =   9600
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Label lblRecorder 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recorder"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   2715
         TabIndex        =   16
         Top             =   3000
         Width           =   1395
      End
      Begin VB.Label lblPouch 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pouch"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   3000
         TabIndex        =   15
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblTrophy 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trophy"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   3000
         TabIndex        =   14
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label lblBroom 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Broom Stick"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Top             =   1560
         Width           =   1785
      End
      Begin VB.Label lblRat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rat"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   3240
         TabIndex        =   12
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblBooks 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pile of Books"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label lblHourGlass 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hour Glass"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   3480
         Width           =   1635
      End
      Begin VB.Label lblFrame 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblPhone 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label lblBasket 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basket"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   3000
         Width           =   1005
      End
      Begin VB.Label lblClock 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wall Clock"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2520
         Width           =   1635
      End
      Begin VB.Label lblBall 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ball"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   810
         TabIndex        =   5
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label lblPen 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pen"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label lblApple 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FF80FF&
         BackStyle       =   0  'Transparent
         Caption         =   "2 Apples"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   465
         TabIndex        =   3
         Top             =   600
         Width           =   1365
      End
   End
   Begin VB.Image imgAppleTwo 
      Height          =   915
      Left            =   8640
      Picture         =   "Form2.frx":1D2EF
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   900
   End
   Begin VB.Image imgRecorder 
      Appearance      =   0  'Flat
      Height          =   600
      Left            =   5520
      Picture         =   "Form2.frx":76936
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1035
   End
   Begin VB.Image imgPouch 
      Appearance      =   0  'Flat
      Height          =   1125
      Left            =   480
      Picture         =   "Form2.frx":946A8
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   1155
   End
   Begin VB.Image imgTrophy 
      Appearance      =   0  'Flat
      Height          =   1440
      Left            =   3840
      Picture         =   "Form2.frx":B2C29
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   795
   End
   Begin VB.Image imgBroom 
      Appearance      =   0  'Flat
      Height          =   4020
      Left            =   -240
      Picture         =   "Form2.frx":C6B16
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1755
   End
   Begin VB.Image imgRat 
      Height          =   825
      Left            =   2760
      Picture         =   "Form2.frx":DFEDE
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgBooks 
      Appearance      =   0  'Flat
      Height          =   1245
      Left            =   2400
      Picture         =   "Form2.frx":1CF247
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1545
   End
   Begin VB.Image imgHourGlass 
      Appearance      =   0  'Flat
      Height          =   1905
      Left            =   7560
      Picture         =   "Form2.frx":211F7D
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   945
   End
   Begin VB.Image imgCp 
      Appearance      =   0  'Flat
      Height          =   780
      Left            =   5880
      Picture         =   "Form2.frx":2A4C13
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   660
   End
   Begin VB.Image imgBall 
      Appearance      =   0  'Flat
      Height          =   1275
      Left            =   3960
      Picture         =   "Form2.frx":2B1C85
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1560
   End
   Begin VB.Image imgBasket 
      Appearance      =   0  'Flat
      Height          =   2115
      Left            =   7560
      Picture         =   "Form2.frx":31BB4E
      Stretch         =   -1  'True
      Top             =   7080
      Width           =   1680
   End
   Begin VB.Image imgPen 
      Appearance      =   0  'Flat
      Height          =   900
      Left            =   480
      Picture         =   "Form2.frx":37D358
      Stretch         =   -1  'True
      Top             =   7920
      Width           =   1020
   End
   Begin VB.Image imgFrame 
      Appearance      =   0  'Flat
      Height          =   1740
      Left            =   240
      Picture         =   "Form2.frx":382A11
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   2010
   End
   Begin VB.Image imgClock 
      Appearance      =   0  'Flat
      Height          =   795
      Left            =   7560
      Picture         =   "Form2.frx":398D7C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   780
   End
   Begin VB.Image imgApple 
      Appearance      =   0  'Flat
      Height          =   675
      Left            =   1560
      Picture         =   "Form2.frx":494CEF
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   600
   End
   Begin VB.Image imgEnd 
      Height          =   735
      Left            =   13560
      Picture         =   "Form2.frx":4EE336
      Stretch         =   -1  'True
      Top             =   8280
      Width           =   735
   End
   Begin VB.Image imgApple2 
      Height          =   9225
      Left            =   0
      Picture         =   "Form2.frx":4F2C51
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9540
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu menuHow 
         Caption         =   "How to Play the Game"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jan 12 2016
'Programmed By: Reyes Nam L.
'This program is A Simple Prototype of Mystery Case Files or Gardenscapes or Alike
'=============================================================
Private Sub cmdBack_Click()
'To show Form 1
Unload Me
Form1.Show
End Sub

Private Sub cmdStart_Click()
Unload Me
Form2.Show
End Sub

Private Sub imgApple_Click()
Me.imgApple.Visible = False
Me.lblApple.Caption = Val(Me.lblApple.Caption) - 1
Me.lblApple.Caption = Me.lblApple.Caption & "Apple"
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgAppleTwo_Click()
Me.imgAppleTwo.Visible = False
Me.lblApple.Caption = Val(Me.lblApple.Caption) - 1
Me.lblApple.Caption = Me.lblApple.Caption & "Apple"
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgBall_Click()
Me.imgBall.Visible = False
Me.lblBall.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgBasket_Click()
Me.imgBasket.Visible = False
Me.lblBasket.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgBooks_Click()
Me.imgBooks.Visible = False
Me.lblBooks.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgBroom_Click()
Me.imgBroom.Visible = False
Me.lblBroom.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgClock_Click()
Me.imgClock.Visible = False
Me.lblClock.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgCp_Click()
Me.imgCp.Visible = False
Me.lblPhone.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgEnd_Click()
'To end the Program
If MsgBox("Are you sure you want to stop playing?", vbQuestion + vbYesNo, "System") = vbYes Then
    End
End If
End Sub

Private Sub imgFrame_Click()
Me.imgFrame.Visible = False
Me.lblFrame.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgHourGlass_Click()
Me.imgHourGlass.Visible = False
Me.lblHourGlass.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgPen_Click()
Me.imgPen.Visible = False
Me.lblPen.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgPouch_Click()
Me.imgPouch.Visible = False
Me.lblPouch.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgRat_Click()
Me.imgRat.Visible = False
Me.lblRat.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub imgRecorder_Click()
Me.imgRecorder.Visible = False
Me.lblRecorder.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub


Private Sub imgTrophy_Click()
Me.imgTrophy.Visible = False
Me.lblTrophy.Visible = False
Me.lblScore.Caption = Val(Me.lblScore.Caption) + 1
End Sub

Private Sub lblApple_Change()
If Me.lblApple.Caption = "0Apple" Then
Me.lblApple.Visible = False
End If
End Sub

Private Sub lblScore_Change()
If Val(Me.lblScore.Caption) = 15 Then
 MsgBox "Good Job!"
Exit Sub
End If
End Sub

Private Sub menuHow_Click()
Unload Me
Form6.Show
End Sub

Private Sub Timer1_Timer()
If Me.lblApple.Caption = "1 Apples" Then
Me.lblApple.Caption = "1 Apple"
End If
End Sub
