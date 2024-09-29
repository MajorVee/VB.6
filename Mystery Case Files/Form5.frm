VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   10020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15210
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   10020
   ScaleMode       =   0  'User
   ScaleWidth      =   15210
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "WhoAmI?"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   6735
      Left            =   10920
      TabIndex        =   1
      Top             =   240
      Width           =   4215
      Begin VB.Image Image1 
         Height          =   3060
         Left            =   240
         Picture         =   "Form5.frx":5C2F5
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   3405
      End
      Begin VB.Image Image2 
         Height          =   3135
         Left            =   1080
         Picture         =   "Form5.frx":606CC
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3060
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reyes, Arvie L. BSCS-I :D"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0FF&
         Height          =   345
         Left            =   480
         TabIndex        =   2
         Top             =   6240
         Width           =   3285
      End
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
      Left            =   120
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      UseMaskColor    =   -1  'True
      Width           =   3135
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Unload Me
Form1.Show
End Sub
