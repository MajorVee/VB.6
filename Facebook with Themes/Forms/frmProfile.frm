VERSION 5.00
Begin VB.Form frmProfile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facebook Profile"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15540
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmProfile.frx":0000
   ScaleHeight     =   9840
   ScaleWidth      =   15540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command10 
      Caption         =   "Post"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11760
      TabIndex        =   13
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   9720
      TabIndex        =   11
      Top             =   4920
      Width           =   1935
      Begin VB.CommandButton Command9 
         Caption         =   "People"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   600
         TabIndex        =   12
         Top             =   0
         Width           =   1335
      End
      Begin VB.Image Image10 
         Height          =   735
         Left            =   -120
         Picture         =   "frmProfile.frx":35A83
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   795
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Franklin Gothic Book"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5400
      MaskColor       =   &H00C0C0FF&
      Picture         =   "frmProfile.frx":367DD
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Friends"
      BeginProperty Font 
         Name            =   "Franklin Gothic Book"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6840
      MaskColor       =   &H00C0C0FF&
      Picture         =   "frmProfile.frx":36F53
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Photos "
      BeginProperty Font 
         Name            =   "Franklin Gothic Book"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8280
      MaskColor       =   &H00C0C0FF&
      Picture         =   "frmProfile.frx":376C9
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "More >"
      BeginProperty Font 
         Name            =   "Franklin Gothic Book"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9720
      MaskColor       =   &H00C0C0FF&
      Picture         =   "frmProfile.frx":37E3F
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H000000FF&
         Caption         =   "X"
         Height          =   615
         Left            =   14640
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   -120
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   14640
         MaskColor       =   &H00C0C0FF&
         Picture         =   "frmProfile.frx":385B5
         TabIndex        =   6
         Top             =   1800
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Activity Log"
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   12840
         MaskColor       =   &H00C0C0FF&
         Picture         =   "frmProfile.frx":38D2B
         TabIndex        =   5
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update Info"
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   11040
         MaskColor       =   &H00C0C0FF&
         Picture         =   "frmProfile.frx":394A1
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Timeline"
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3840
         MaskColor       =   &H00C0C0FF&
         Picture         =   "frmProfile.frx":39C17
         TabIndex        =   1
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Nam T. Carbonel)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   405
         Left            =   3960
         TabIndex        =   3
         Top             =   1800
         Width           =   2985
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jean L. Reyes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Left            =   4080
         TabIndex        =   2
         Top             =   1320
         Width           =   2745
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   2775
         Left            =   120
         Picture         =   "frmProfile.frx":3A38D
         Stretch         =   -1  'True
         Top             =   120
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   3795
         Left            =   0
         Picture         =   "frmProfile.frx":4D732
         Stretch         =   -1  'True
         Top             =   -1080
         Width           =   15480
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jean L. Reyes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   6120
      TabIndex        =   15
      Top             =   6000
      Width           =   1500
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   5400
      Picture         =   "frmProfile.frx":9014C
      Stretch         =   -1  'True
      Top             =   8760
      Width           =   495
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   5280
      Picture         =   "frmProfile.frx":A34F1
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   735
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   5400
      Picture         =   "frmProfile.frx":B6896
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   615
   End
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
If MsgBox("Are you sure you want to close your Facebook Profile?", vbQuestion + vbYesNo, "Facebook") = vbYes Then
    Unload Me
frmThemes.Enabled = True
End If
End Sub

Private Sub Form_Load()
frmThemes.Enabled = False
End Sub
