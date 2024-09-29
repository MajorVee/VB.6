VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Search"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmSearch.frx":0000
   ScaleHeight     =   7830
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   2400
   End
   Begin Project1.jcbutton cmdSearch 
      Height          =   375
      Left            =   10560
      TabIndex        =   16
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ButtonStyle     =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Schoolbook"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "SEARCH"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.ComboBox cboRecord 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   0
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1440
      Width           =   2535
   End
   Begin VB.OptionButton optClient 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Client"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
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
      Top             =   1560
      Width           =   1095
   End
   Begin VB.OptionButton optCar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Car"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   855
   End
   Begin VB.OptionButton optDriver 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Driver"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.ComboBox cboRecord 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   1
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1485
      Width           =   2535
   End
   Begin VB.ComboBox cboRecord 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Index           =   2
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1485
      Width           =   2535
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   4575
      Left            =   240
      TabIndex        =   7
      Top             =   3120
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   12632256
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11400
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":5601
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      ToolTipText     =   "Type your search query and Press Enter"
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label lblS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   6000
      TabIndex        =   21
      Top             =   2400
      Width           =   285
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   600
      Picture         =   "frmSearch.frx":5B9B
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   19
      Top             =   210
      Width           =   1305
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search by:"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   360
      TabIndex        =   18
      Top             =   1100
      Width           =   1380
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   6360
      TabIndex        =   17
      Top             =   1095
      Width           =   990
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Registered Driver(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12675
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   2145
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   13800
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Rented Car(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   13365
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   13800
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Total No. of Car(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   13140
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   13800
      TabIndex        =   10
      Top             =   4800
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   630
      Left            =   13680
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Registered Client(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   12810
      TabIndex        =   8
      Top             =   3480
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000008&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   240
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000008&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   6240
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   0
      TabIndex        =   20
      Top             =   120
      Width           =   20655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   2520
      TabIndex        =   22
      Top             =   2280
      Width           =   7575
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSearch_Click()
If optClient.Value = True And cboRecord(0).ListIndex >= 0 Then
        Set rs = New ADODB.Recordset
            With rs
                    rs.Open "Select * from Client where " & cboRecord(0).Text & " like'%" & txtSearch.Text & "%'", conn, 3, 3
                    If .RecordCount >= 1 Then
                        lvSearch.ListItems.Clear
                        Do While Not .EOF
                            lvSearch.ListItems.Add , , !ClientID, 1, 1
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(1) = "" & !FName
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(2) = "" & !MName
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(3) = "" & !LName
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(4) = "" & !Address
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(5) = "" & !DOB
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(6) = "" & !Gender
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(7) = "" & !Occup
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(8) = "" & !PhoneNo
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(9) = "" & !Email
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(10) = "" & !License
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(11) = "" & !DateReg
                            'Label4.Caption = "Total Record(s)  : " & lvFilter.ListItems.Count
                            .MoveNext
                        Loop
                    Else
                        MsgBox "No Record Found!", vbExclamation, "Warning"
                        txtSearch.Text = ""
                        txtSearch.SetFocus
                        Exit Sub
                    End If
                    .Close
            End With
    ElseIf optCar.Value = True And cboRecord(1).ListIndex >= 0 Then
        Set rs = New ADODB.Recordset
            With rs
                    rs.Open "Select * from Car where " & cboRecord(1).Text & " like'%" & txtSearch.Text & "%'", conn, 3, 3
                    If .RecordCount >= 1 Then
                        lvSearch.ListItems.Clear
                        Do While Not .EOF
                            lvSearch.ListItems.Add , , !CarID, 1, 1
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(1) = "" & !PlateNo
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(2) = "" & !RegNo
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(3) = "" & !Type
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(4) = "" & !Trip
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(5) = "" & !DateManu
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(6) = "" & !Model
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(7) = "" & !Make
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(8) = "" & !Speed
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(9) = "" & !Condition
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(10) = "" & !AvailStatus
                            'Label4.Caption = "Total Record(s)  : " & lvFilter.ListItems.Count
                            .MoveNext
                        Loop
                    Else
                        MsgBox "No Record Found!", vbExclamation, "Warning"
                        txtSearch.Text = ""
                        txtSearch.SetFocus
                        Exit Sub
                    End If
                    .Close
            End With
    ElseIf optDriver.Value = True And cboRecord(2).ListIndex >= 0 Then
        Set rs = New ADODB.Recordset
            With rs
                    rs.Open "Select * from Driver where " & cboRecord(2).Text & " like'%" & txtSearch.Text & "%'", conn, 3, 3
                    If .RecordCount >= 1 Then
                        lvSearch.ListItems.Clear
                        Do While Not .EOF
                            lvSearch.ListItems.Add , , !DriverID, 1, 1
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(1) = "" & !DriverName
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(2) = "" & !Address
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(3) = "" & !BDate
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(4) = "" & !PhoneNo
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(5) = "" & !License
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(6) = "" & !DateReg
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(7) = "" & !CStatus
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(8) = "" & !TIN
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(9) = "" & !SSS
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(10) = "" & !PlateNo
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(11) = "" & !Type
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(12) = "" & !Make
                            lvSearch.ListItems(lvSearch.ListItems.Count).SubItems(13) = "" & !Trip
                            'Label4.Caption = "Total Record(s)  : " & lvFilter.ListItems.Count
                            .MoveNext
                        Loop
                    Else
                        MsgBox "No Record Found!", vbExclamation, "Warning"
                        txtSearch.Text = ""
                        lvSearch.ListItems.Clear
                        txtSearch.SetFocus
                        Exit Sub
                    End If
                    .Close
            End With
    
    Else
        MsgBox "Please select a record to filter the data.", vbExclamation, "Try Again!"
        txtSearch.SetFocus
    End If
End Sub

Private Sub Form_Load()
modConnect.Connected

Call NoClientSearch
Call NoCarsSearch
Call NoDriverSearch
Call NoCarUnavail
Call NoRentedSearch

cboRecord(0).AddItem "FName"
cboRecord(0).AddItem "LName"
cboRecord(0).AddItem "License"
cboRecord(0).AddItem "PhoneNo"

cboRecord(1).AddItem "PlateNo"
'cboRecord(1).AddItem "CarID"
cboRecord(1).AddItem "Trip"
cboRecord(1).AddItem "Model"
cboRecord(1).AddItem "AvailStatus"

cboRecord(2).AddItem "DriverName"
cboRecord(2).AddItem "PhoneNo"
cboRecord(2).AddItem "License"
cboRecord(2).AddItem "DriverID"
End Sub
Private Sub loadLVDriver()
With lvSearch
    .View = lvwReport
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Driver ID", 1600
    .ColumnHeaders.Add , , "Name", 1700
    .ColumnHeaders.Add , , "Address", 1700
    .ColumnHeaders.Add , , "Birth", 1700
    .ColumnHeaders.Add , , "Phone", 2000
    .ColumnHeaders.Add , , "License", 1500
    .ColumnHeaders.Add , , "Date Registered", 2000
    .ColumnHeaders.Add , , "Civil Status", 1500
    .ColumnHeaders.Add , , "TIN", 0
    .ColumnHeaders.Add , , "SSS", 0
    .ColumnHeaders.Add , , "Plate No.", 1800
    .ColumnHeaders.Add , , "Type", 0
    .ColumnHeaders.Add , , "Make", 0
    .ColumnHeaders.Add , , "Trip", 0
End With
End Sub

Private Sub loadLVCLient()
With lvSearch
    .View = lvwReport
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Client ID", 1200
    .ColumnHeaders.Add , , "Firstname", 1700
    .ColumnHeaders.Add , , "Midname", 1700
    .ColumnHeaders.Add , , "Lastname", 1700
    .ColumnHeaders.Add , , "Address", 2000
    .ColumnHeaders.Add , , "Birth", 1500
    .ColumnHeaders.Add , , "Gender", 1300
    .ColumnHeaders.Add , , "Occupation", 1800
    .ColumnHeaders.Add , , "Phone", 1800
    .ColumnHeaders.Add , , "Email", 2300
    .ColumnHeaders.Add , , "License", 2000
    .ColumnHeaders.Add , , "Date Registered", 2000
    
End With
End Sub
Private Sub loadLVCar()
With lvSearch
    .View = lvwReport
    .ColumnHeaders.Clear
    .ColumnHeaders.Add , , "Car ID", 0
    .ColumnHeaders.Add , , "Plate No", 2300
    .ColumnHeaders.Add , , "Registration No", 2800
    .ColumnHeaders.Add , , "Vehicle Type", 0
    .ColumnHeaders.Add , , "Trip", 800
    .ColumnHeaders.Add , , "Date Manufactured", 0
    .ColumnHeaders.Add , , "Model", 1400
    .ColumnHeaders.Add , , "Make", 0
    .ColumnHeaders.Add , , "Speed", 1200
    .ColumnHeaders.Add , , "Condition", 1500
    .ColumnHeaders.Add , , "Status", 1600
End With
End Sub

Private Sub optCar_Click()
cboRecord(0).Visible = False
cboRecord(1).Visible = True
cboRecord(2).Visible = False
Call loadLVCar
If optClient.Value = False And optDriver.Value = False Then
    lvSearch.ListItems.Clear
    txtSearch.Text = ""
End If
End Sub

Private Sub optClient_Click()
cboRecord(0).Visible = True
cboRecord(1).Visible = False
cboRecord(2).Visible = False
Call loadLVCLient
If optCar.Value = False And optDriver.Value = False Then
    lvSearch.ListItems.Clear
    txtSearch.Text = ""
End If
End Sub

Private Sub optDriver_Click()
cboRecord(0).Visible = False
cboRecord(1).Visible = False
cboRecord(2).Visible = True
Call loadLVDriver
If optClient.Value = False And optCar.Value = False Then
    lvSearch.ListItems.Clear
    txtSearch.Text = ""
End If
End Sub

Private Sub Timer1_Timer()
If val(Me.lvSearch.ListItems.Count) > 1 Then
Me.lblS.Caption = "There are " & Me.lvSearch.ListItems.Count & " items found in the list."
ElseIf Me.lvSearch.ListItems.Count = 1 Then
Me.lblS.Caption = "There is " & Me.lvSearch.ListItems.Count & " item found in the list."
Else
Me.lblS.Caption = "No items found in the list."
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdSearch_Click
Else
End If
End Sub
