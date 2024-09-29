VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Van Rental Management System"
   ClientHeight    =   8745
   ClientLeft      =   240
   ClientTop       =   840
   ClientWidth     =   16695
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   16695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrLoad 
      Interval        =   1
      Left            =   1200
      Top             =   1080
   End
   Begin Project1.CommandLive cmdDrivers 
      Height          =   1830
      Left            =   5880
      TabIndex        =   22
      Top             =   2160
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3228
      Picture         =   "frmMain.frx":BC3F
      Animate         =   -1  'True
      Caption         =   "Drivers"
      Detail          =   "Label1"
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BackColor       =   4194368
   End
   Begin Project1.CommandLive cmdUA 
      Height          =   1830
      Left            =   1440
      TabIndex        =   14
      Top             =   2160
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3228
      Picture         =   "frmMain.frx":101DF
      Caption         =   "User Accounts"
      Detail          =   ""
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BackColor       =   8421504
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   15840
      Top             =   3480
   End
   Begin Project1.jcbutton cmdLogout 
      Height          =   615
      Left            =   16440
      TabIndex        =   2
      Top             =   9360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      Caption         =   "Log - Out"
      UseMaskCOlor    =   -1  'True
   End
   Begin Project1.CommandLive cmdClient 
      Height          =   1830
      Left            =   1440
      TabIndex        =   15
      Top             =   4200
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3228
      Picture         =   "frmMain.frx":15D1C
      Caption         =   "Clients"
      Detail          =   "Label1"
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BackColor       =   65280
   End
   Begin Project1.CommandLive cmdRent 
      Height          =   1830
      Left            =   1440
      TabIndex        =   16
      Top             =   6240
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3228
      Picture         =   "frmMain.frx":1A6E1
      Caption         =   "Rent Car"
      Detail          =   "Label1"
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BackColor       =   4210688
   End
   Begin Project1.CommandLive cmdLog 
      Height          =   1830
      Left            =   5880
      TabIndex        =   17
      Top             =   4200
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3228
      Picture         =   "frmMain.frx":1EB29
      Caption         =   "Log History"
      Detail          =   "Label1"
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BackColor       =   4210816
   End
   Begin Project1.CommandLive cmdReturn 
      Height          =   1830
      Left            =   5880
      TabIndex        =   18
      Top             =   6240
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3228
      Picture         =   "frmMain.frx":23D26
      Caption         =   "Return Vehicle"
      Detail          =   "Label1"
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BackColor       =   -2147483646
   End
   Begin Project1.CommandLive cmdSearch 
      Height          =   1830
      Left            =   10320
      TabIndex        =   19
      Top             =   4200
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3228
      Picture         =   "frmMain.frx":28802
      Caption         =   "Search"
      Detail          =   "Label1"
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BackColor       =   0
   End
   Begin Project1.CommandLive cmdCars 
      Height          =   1830
      Left            =   10320
      TabIndex        =   20
      Top             =   2160
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3228
      Picture         =   "frmMain.frx":2DC2B
      Caption         =   "Vans"
      Detail          =   "Label1"
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BackColor       =   49344
   End
   Begin Project1.CommandLive cmdReports 
      Height          =   1830
      Left            =   10320
      TabIndex        =   21
      Top             =   6240
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   3228
      Picture         =   "frmMain.frx":32DE7
      Caption         =   "Reports"
      Detail          =   "Label1"
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BorderColor     =   -2147483635
      BackColor       =   16576
   End
   Begin Project1.jcbutton cmdClose 
      Height          =   615
      Left            =   18000
      TabIndex        =   25
      Top             =   9360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      Caption         =   "Close System"
      UseMaskCOlor    =   -1  'True
   End
   Begin Project1.jcbutton cmdAbout 
      Height          =   375
      Left            =   16560
      TabIndex        =   26
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      Caption         =   "About"
      UseMaskCOlor    =   -1  'True
   End
   Begin Project1.jcbutton cmdbackup 
      Height          =   495
      Left            =   13800
      TabIndex        =   27
      Top             =   9240
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12632256
      Caption         =   "Backup Database"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   495
      Left            =   15600
      TabIndex        =   24
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "11:11 PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   18240
      TabIndex        =   23
      Top             =   7200
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "5:47:20 AM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   17160
      TabIndex        =   13
      Top             =   3840
      Width           =   2475
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "5:47:20 AM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   585
      Left            =   17160
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Image Image1 
      Height          =   945
      Left            =   16080
      Picture         =   "frmMain.frx":376D4
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   17280
      TabIndex        =   11
      Top             =   4320
      Width           =   2190
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "August 28, 2017"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   17025
      TabIndex        =   10
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ADMINISTRATOR"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   16485
      TabIndex        =   9
      Top             =   6120
      Width           =   3240
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome:"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   16200
      TabIndex        =   8
      Top             =   5520
      Width           =   2040
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nam L. Reyes"
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   18150
      TabIndex        =   7
      Top             =   6840
      Width           =   1470
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Current User:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   16320
      TabIndex        =   6
      Top             =   6840
      Width           =   1425
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Time Logged In:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   16320
      TabIndex        =   5
      Top             =   7200
      Width           =   1710
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "08/28/2017"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   18240
      TabIndex        =   4
      Top             =   7560
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Logged In:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   16320
      TabIndex        =   3
      Top             =   7560
      Width           =   1725
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   15960
      TabIndex        =   1
      Top             =   5400
      Width           =   3975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1515
      Left            =   15990
      TabIndex        =   0
      Top             =   3720
      Width           =   3915
   End
   Begin VB.Menu mnuOp 
      Caption         =   "Options"
      Begin VB.Menu mnuBD 
         Caption         =   "Backup Database"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAbout_Click()
frmAbout.Show 1
End Sub

Private Sub cmdbackup_Click()
frmBackup.Show 1

End Sub

Private Sub cmdCars_Click()
If Me.Label9.Caption = "GUEST" Then
frmError.Show 1
Else
frmCar.Show 1
End If
End Sub

Private Sub cmdClient_Click()
If Me.Label9.Caption = "GUEST" Then
frmError.Show 1
Else
frmClient.Show 1
End If
End Sub

Private Sub cmdClose_Click()
Dim logs As String
        modConnect.Connected
        Set rs = New ADODB.Recordset
        logs = Label8.Caption
        rs.Open "Select * from Logs where LogID=" & logs & "", conn, adOpenDynamic, adLockOptimistic
            With rs
                !TimeOut = Time
                .Update
            End With
        rs.Close
        Set rs = Nothing
End
End Sub

Private Sub cmdDrivers_Click()
If Me.Label9.Caption = "OPERATOR" Then
frmError.Show 1
ElseIf Me.Label9.Caption = "GUEST" Then
frmError.Show 1
Else
frmDrivers.Show 1
End If
End Sub

Private Sub cmdLog_Click()
frmLog.Show 1
End Sub

Private Sub cmdLogout_Click()
 
     Dim logs As String
        modConnect.Connected
        Set rs = New ADODB.Recordset
        logs = Label8.Caption
        rs.Open "Select * from Logs where LogID=" & logs & "", conn, adOpenDynamic, adLockOptimistic
            With rs
                !TimeOut = Time
                .Update
            End With
        rs.Close
        Set rs = Nothing
        
If MsgBox("Are you sure you want to log-out?", vbQuestion + vbYesNo) = vbYes Then
' Unload Me
 frmLock.Show 1
End If

End Sub

Private Sub cmdRent_Click()
If Me.Label9.Caption = "GUEST" Then
frmError.Show 1
Else
frmRent.Show 1
End If
End Sub

Private Sub cmdReports_Click()
If Me.Label9.Caption = "GUEST" Then
frmMain.mnuBD.Enabled = True
frmError.Show 1
Else
frmReports.Show 1
End If

End Sub

Private Sub cmdReturn_Click()
If Me.Label9.Caption = "GUEST" Then
frmError.Show 1
Else
frmReturn.Show 1
End If
End Sub

Private Sub cmdSearch_Click()
frmSearch.Show 1
End Sub

Private Sub cmdUA_Click()
If Me.Label9.Caption = "OPERATOR" Then
frmError.Show 1
ElseIf Me.Label9.Caption = "GUEST" Then
frmError.Show 1
Else
frmUser.Show 1
End If
End Sub

Private Sub Form_Load()

Label6.Caption = WeekdayName(Weekday(Now))
Label7.Caption = MonthName(Month(Now)) & " " & Day(Now) & ", " & Year(Now)
Label4.Caption = DateValue(Now)
'Label12.Caption = Time

Label15.Caption = DateValue(Now)

Me.cmdUA.Animate = True
Me.cmdUA.Interval = 250
Me.cmdUA.Detail = " List of registered User Accounts."

Me.cmdClient.Animate = True
Me.cmdClient.Interval = 450
Me.cmdClient.Detail = " List of registered clients."

Me.cmdRent.Animate = True
Me.cmdRent.Interval = 650
Me.cmdRent.Detail = "To rent a Vehicle."

Me.cmdDrivers.Animate = True
Me.cmdDrivers.Interval = 850
Me.cmdDrivers.Detail = " List of registered drivers."

Me.cmdLog.Animate = True
Me.cmdLog.Interval = 1050
Me.cmdLog.Detail = " Shows the time of log-in and log-out of a user."

Me.cmdReturn.Animate = True
Me.cmdReturn.Interval = 1250
Me.cmdReturn.Detail = " To return a van."

Me.cmdCars.Animate = True
Me.cmdCars.Interval = 1450
Me.cmdCars.Detail = " List of the available and rented van. It also shows information about the van."

Me.cmdSearch.Animate = True
Me.cmdSearch.Interval = 1650
Me.cmdSearch.Detail = " Search for users, clients, drivers, available van."

Me.cmdReports.Animate = True
Me.cmdReports.Interval = 1850
Me.cmdReports.Detail = " Summary of all vans, drivers, employees, transactions..."

End Sub

Private Sub mnuAbout_Click()
cmdAbout_Click
End Sub

Private Sub mnuBD_Click()
Call cmdbackup_Click
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Time
End Sub

Private Sub tmrLoad_Timer()
frmLock.Show 1
End Sub

