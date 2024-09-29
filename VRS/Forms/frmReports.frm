VERSION 5.00
Begin VB.Form frmReports 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reports"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.CommandLive cmdUA 
      Height          =   1830
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   3750
      _extentx        =   6615
      _extenty        =   3228
      caption         =   "User Accounts Masterlist"
      detail          =   ""
      backcolor       =   8421504
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      picture         =   "frmReports.frx":0000
   End
   Begin Project1.CommandLive cmdDrivers 
      Height          =   1830
      Left            =   4080
      TabIndex        =   1
      Top             =   1560
      Width           =   3750
      _extentx        =   6615
      _extenty        =   3228
      caption         =   "Drivers Masterlist"
      detail          =   "Label1"
      animate         =   -1  'True
      backcolor       =   4194368
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      picture         =   "frmReports.frx":5B3E
   End
   Begin Project1.CommandLive cmdClient 
      Height          =   1830
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   3750
      _extentx        =   6615
      _extenty        =   3228
      caption         =   "Clients Masterlist"
      detail          =   "Label1"
      backcolor       =   65280
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      picture         =   "frmReports.frx":A0E0
   End
   Begin Project1.CommandLive cmdCars 
      Height          =   1830
      Left            =   8040
      TabIndex        =   3
      Top             =   1560
      Width           =   3750
      _extentx        =   6615
      _extenty        =   3228
      caption         =   "Vans Masterlist"
      detail          =   "Label1"
      backcolor       =   49344
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      picture         =   "frmReports.frx":EAA6
   End
   Begin Project1.CommandLive cmdReturn 
      Height          =   1830
      Left            =   8040
      TabIndex        =   6
      Top             =   3600
      Width           =   3750
      _extentx        =   6615
      _extenty        =   3228
      caption         =   "Vans to be returned"
      detail          =   "Label1"
      backcolor       =   -2147483646
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      picture         =   "frmReports.frx":13C64
   End
   Begin Project1.CommandLive cmdRent 
      Height          =   1830
      Left            =   4080
      TabIndex        =   7
      Top             =   3600
      Width           =   3750
      _extentx        =   6615
      _extenty        =   3228
      caption         =   "Rented Vans Records"
      detail          =   "Label1"
      backcolor       =   4210688
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      picture         =   "frmReports.frx":18742
   End
   Begin Project1.CommandLive cmdRCR 
      Height          =   1830
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   3750
      _extentx        =   6615
      _extenty        =   3228
      caption         =   "Returned Car Report"
      detail          =   "Label1"
      backcolor       =   -2147483646
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      bordercolor     =   -2147483635
      picture         =   "frmReports.frx":1CB8C
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   720
      Picture         =   "frmReports.frx":21D3E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Reports"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   1920
      TabIndex        =   4
      Top             =   315
      Width           =   1710
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   20655
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCars_Click()
Set rs = New ADODB.Recordset
        rs.Open "Select * from Car", conn, 3, 3
        Set rptCars.DataSource = rs
        Set rs = Nothing
        rptCars.Show 1
End Sub

Private Sub cmdClient_Click()
     Set rs = New ADODB.Recordset
        rs.Open "Select * from Client", conn, 3, 3
        Set rptClient.DataSource = rs
        Set rs = Nothing
        rptClient.Show 1
End Sub

Private Sub cmdDrivers_Click()
 Set rs = New ADODB.Recordset
        rs.Open "Select * from Driver", conn, 3, 3
        Set rptDriverList.DataSource = rs
        Set rs = Nothing
        rptDriverList.Show 1
End Sub

Private Sub cmdRCR_Click()
Set rs = New ADODB.Recordset
        rs.Open "Select * from Purchase_Order", conn, 3, 3
        Set rptRented.DataSource = rs
        Set rs = Nothing
        rptRented.Show 1
End Sub

Private Sub cmdRent_Click()
   Set rs = New ADODB.Recordset
        rs.Open "Select * from Order_Transaction", conn, 3, 3
        Set rptTransList.DataSource = rs
        Set rs = Nothing
        rptTransList.Show 1
End Sub

Private Sub cmdReturn_Click()
Set rs = New ADODB.Recordset
        rs.Open "Select * from Purchase_Order", conn, 3, 3
        Set rptRented.DataSource = rs
        Set rs = Nothing
        rptRented.Show 1
End Sub

Private Sub cmdUA_Click()
Set rs = New ADODB.Recordset
        rs.Open "Select * from [User]", conn, 3, 3
        Set rptUsers.DataSource = rs
        Set rs = Nothing
        rptUsers.Show 1
End Sub
