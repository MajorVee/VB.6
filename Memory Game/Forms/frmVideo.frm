VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "CODEJO~1.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}#4.0#0"; "mshtml.tlb"
Begin VB.Form frmVideo 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Game"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13290
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   13290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1100
      Left            =   13560
      Top             =   1680
   End
   Begin MSHTMLCtl.Scriptlet Scriptlet1 
      Height          =   615
      Left            =   13440
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   615
      Scrollbar       =   0   'False
      URL             =   "about:blank"
   End
   Begin VB.Label lblClickHere 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "P L A Y"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Top             =   8400
      Width           =   2295
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   13440
      Top             =   2400
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpMoviePreview 
      Height          =   10335
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13335
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   23521
      _cy             =   18230
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer
Dim a As Integer
Private Sub Form_Load()
Call LoadMoviePreview
Me.lblClickHere.Visible = False
a = 1
End Sub

Private Sub lblClickHere_Click()
Unload Me
frmMain.Show
End Sub

Private Sub lblClickHere_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.lblClickHere.BackColor = &HFF80FF
Me.lblClickHere.ForeColor = vbYellow
End Sub

Private Sub Timer1_Timer()
a = a + 1
If a = 7 Then Me.lblClickHere.Visible = True
End Sub
