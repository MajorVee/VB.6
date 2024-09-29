VERSION 5.00
Begin VB.Form Dialog 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kape-Libro Coffee Shop"
   ClientHeight    =   3075
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Dialog.frx":0000
   ScaleHeight     =   3075
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   3915
      Left            =   -120
      Picture         =   "Dialog.frx":9A40
      Stretch         =   -1  'True
      Top             =   -1080
      Width           =   1500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Programmed by: Sardua -.- and  Reyes :}"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   4635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   """Kape-Libro Coffee Shop"""
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   4920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic POS without Database"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   3885
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
Unload Me
frmOpen.Show
End Sub
