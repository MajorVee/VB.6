VERSION 5.00
Begin VB.Form frmError 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Error"
   ClientHeight    =   1575
   ClientLeft      =   2760
   ClientTop       =   3690
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.jcbutton jcbutton1 
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "CLOSE"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Access Denied!"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   555
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   3450
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   360
      OLEDropMode     =   1  'Manual
      Picture         =   "frmError.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1140
   End
End
Attribute VB_Name = "frmError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub jcbutton1_Click()
Unload Me
End Sub
