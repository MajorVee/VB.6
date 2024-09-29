VERSION 5.00
Begin VB.Form frmReturn 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Return Vehicle"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   17985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmReturn.frx":0000
   ScaleHeight     =   6105
   ScaleWidth      =   17985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.jcbutton cmdReturn 
      Height          =   615
      Left            =   14640
      TabIndex        =   27
      Top             =   5040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      ButtonStyle     =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "RETURN"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.TextBox txtCar 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   420
      Left            =   10800
      TabIndex        =   23
      Top             =   3240
      Width           =   6855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "Van Information"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   3015
      Left            =   9000
      TabIndex        =   15
      Top             =   1440
      Width           =   8895
      Begin VB.TextBox txtCarID 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   420
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   600
         Width           =   6855
      End
      Begin VB.TextBox txtTank 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   420
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   3720
         Width           =   6855
      End
      Begin VB.TextBox txtMil 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   420
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   3120
         Width           =   6855
      End
      Begin VB.TextBox txtCond 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   420
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2400
         Width           =   6855
      End
      Begin VB.TextBox txtPlate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   420
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1200
         Width           =   6855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver's ID:"
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
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   1185
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tank Level:"
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
         Left            =   240
         TabIndex        =   20
         Top             =   3720
         Width           =   1275
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mileage:"
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
         Left            =   240
         TabIndex        =   19
         Top             =   3120
         Width           =   1005
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Condition:"
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
         Left            =   240
         TabIndex        =   18
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Car Model:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1800
         Width           =   1290
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate No:"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Rent Information"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   8775
      Begin VB.ComboBox cboClient 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   6495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
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
         Left            =   4080
         TabIndex        =   30
         Top             =   2520
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
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
         Left            =   1080
         TabIndex        =   29
         Top             =   2520
         Width           =   630
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Day(s) :"
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
         Left            =   6600
         TabIndex        =   14
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trip from:"
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
         Left            =   720
         TabIndex        =   13
         Top             =   3120
         Width           =   1050
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Name:"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver's Name:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID:"
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
         Left            =   720
         TabIndex        =   9
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label lblRet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   8
         Top             =   1200
         Width           =   6495
      End
      Begin VB.Label lblRet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   1800
         Width           =   6495
      End
      Begin VB.Label lblRet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   6
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblRet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4560
         TabIndex        =   5
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblRet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   7560
         TabIndex        =   4
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblRet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   2040
         TabIndex        =   3
         Top             =   3120
         Width           =   6495
      End
   End
   Begin Project1.jcbutton cmdClear 
      Height          =   615
      Left            =   11520
      TabIndex        =   32
      Top             =   5040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      ButtonStyle     =   9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "C L E A R"
      UseMaskCOlor    =   -1  'True
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Van Return"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   285
      Width           =   2625
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   480
      Picture         =   "frmReturn.frx":5601
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   0
      TabIndex        =   31
      Top             =   240
      Width           =   20655
   End
End
Attribute VB_Name = "frmReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboClient_Click()
Call DisplayClient
Call DisplayDriver
End Sub

Public Sub DisplayDriver()
On Error Resume Next
Set rs = New ADODB.Recordset
rs.Open "Select * from Rent_Car where DriverID='" & Me.txtCarID.Text & "'", conn, 3, 3
txtPlate.Text = rs("PlateNo") 'debug jud sya sugod dri kasi walai unod ang drivers id maong dli mu load
txtCar.Text = rs("Model")
txtCond.Text = rs("Condition")
txtMil.Text = rs("MileAge")
txtTank.Text = rs("TankLevel")
Set rs = Nothing
End Sub

Public Sub DisplayClient()
Set rs = New ADODB.Recordset
rs.Open "Select * from Purchase_Order where ClientID='" & cboClient.Text & "'", conn, 3, 3
lblRet(0).Caption = rs!ClientName
lblRet(1).Caption = rs!DriverName
lblRet(2).Caption = rs!DateFrom
lblRet(3).Caption = rs!DateTo
lblRet(4).Caption = rs!NoDay
lblRet(5).Caption = rs!TripTo
Me.txtCarID.Text = rs!DriverID
Set rs = Nothing
End Sub

Public Sub PopClientID()
Set rs = New ADODB.Recordset
    rs.Open "Select ClientID from Purchase_Order Order by ClientID ASC", conn, 3, 3
        Do While Not rs.EOF
            cboClient.AddItem rs!ClientID
            rs.MoveNext
        Loop
    Set rs = Nothing
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClear_Click()
cboClient.Clear
lblRet(0).Caption = ""
lblRet(1).Caption = ""
lblRet(2).Caption = ""
lblRet(3).Caption = ""
lblRet(4).Caption = ""
lblRet(5).Caption = ""
txtCarID.Text = ""
Me.txtPlate.Text = ""
Me.txtCar.Text = ""
Me.txtCond.Text = ""
Me.txtMil.Text = ""
Me.txtTank.Text = ""

Call PopClientID
End Sub

Private Sub cmdReturn_Click()
     If cboClient.Text = vbNullString And txtCarID.Text = vbNullString Then
            MsgBox "Please select the corresponding id in the list.", vbExclamation, "Warning!"
        Else
            Set rs4 = New ADODB.Recordset
            rs4.Open "Select * from Return_Car", conn, adOpenKeyset, adLockPessimistic
                With rs4
                    .AddNew
                    .Fields("ClientName") = cboClient.Text
                    .Fields("DriverName") = lblRet(1).Caption
                    .Fields("PlateNo") = txtPlate.Text
                    .Fields("Model") = txtCar.Text
                    .Fields("DateReturned") = Date
                    .Update
                End With
                Set rs4 = Nothing

            Set rs = New ADODB.Recordset
            rs.Open "Select * from Purchase_Order where ClientID='" & cboClient.Text & "'", conn, adOpenKeyset, adLockPessimistic
            With rs
                .Delete
                .Close
            End With
            Set rs = Nothing
                        
            Set rs2 = New ADODB.Recordset
            rs2.Open "Select AvailStatus, PlateNo from Car where PlateNo='" & Me.txtPlate.Text & "'", conn, 3, 3
               With rs2
                    !AvailStatus = "Available"
                    .Update
                End With
            Set rs2 = Nothing
            
            Set rs3 = New ADODB.Recordset
            rs3.Open "Select * from Rent_Car where DriverID='" & txtCarID.Text & "'", conn, adOpenKeyset, adLockPessimistic
            With rs3
                .Delete
                .Close
            End With
            Set rs3 = Nothing
            

            MsgBox "Van Successfully Returned!", vbInformation, "Success Returned!"
        End If
Call cmdClear_Click
Unload Me

End Sub

Private Sub Form_Load()
modConnect.Connected
Call PopClientID
'Call PopDriverID
End Sub
'Public Sub PopDriverID()
'Set rs = New ADODB.Recordset
'    rs.Open "Select DriverID from Rent_Car Order by DriverID ASC", conn, 3, 3
'        Do While Not rs.EOF
'            cboCarI.AddItem rs!DriverID
'            rs.MoveNext
'        Loop
'    Set rs = Nothing
'End Sub


