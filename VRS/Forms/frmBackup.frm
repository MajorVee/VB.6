VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackup 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Database"
   ClientHeight    =   3255
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.jcbutton btnBackup 
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "Backup Database"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1200
      Top             =   480
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "---------------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2430
      TabIndex        =   3
      Top             =   1200
      Width           =   1155
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "To access the backup database file, please refer to the Backup Folder where the application is installed."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   5940
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Timer2.Enabled = False
End Sub

Private Sub btnBackup_Click()
Call MDbackupdatabases
End Sub

Private Sub Timer2_Timer()
Dim a As Long
Timer2.Enabled = True
ProgressBar1.Max = 101
ProgressBar1.Value = ProgressBar1.Value + 1
btnBackup.Enabled = False
If ProgressBar1.Value = 20 Then
Label1.Caption = "Preparing to Backup Files..."
ElseIf ProgressBar1.Value = 40 Then
Label1.Caption = "Loading Database..."
ElseIf ProgressBar1.Value = 60 Then
Label1.Caption = "Loading Access..."
ElseIf ProgressBar1.Value = 80 Then
Label1.Caption = "Loading Contents..."
ElseIf ProgressBar1.Value = 90 Then
Label1.Caption = "Backup Complete..."
ElseIf ProgressBar1.Value = 101 Then
MsgBox "Backup Completed!!!", vbInformation, "Back Up Done!"
ProgressBar1.Value = 0
Timer2.Enabled = False
Label1.Caption = ""
btnBackup.Enabled = True
Unload Me
End If

End Sub
