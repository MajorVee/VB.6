VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDrivers 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Drivers"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   19980
   BeginProperty Font 
      Name            =   "Century Gothic"
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
   Picture         =   "frmDrivers.frx":0000
   ScaleHeight     =   9000
   ScaleWidth      =   19980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.jcbutton cmdPD 
      Height          =   615
      Left            =   3480
      TabIndex        =   38
      Top             =   6960
      Width           =   3135
      _extentx        =   5530
      _extenty        =   1085
      buttonstyle     =   10
      font            =   "frmDrivers.frx":5601
      backcolor       =   12632256
      caption         =   "Print this Driver's Record"
      usemaskcolor    =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10920
      Top             =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Drivers Information"
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
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   9975
      Begin VB.ComboBox cboMake 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1200
         TabIndex        =   36
         Top             =   5640
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.ComboBox cboColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7080
         TabIndex        =   35
         Top             =   4920
         Width           =   2535
      End
      Begin VB.ComboBox cboType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmDrivers.frx":5631
         Left            =   960
         List            =   "frmDrivers.frx":5633
         TabIndex        =   25
         Top             =   6240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cboPlate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   20
         Top             =   4920
         Width           =   3255
      End
      Begin VB.ComboBox cboCivStat 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   18
         Top             =   3720
         Width           =   7335
      End
      Begin VB.TextBox txtTIN 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   7920
         TabIndex        =   17
         Top             =   4440
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.TextBox txtSSS 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   7920
         TabIndex        =   16
         Top             =   5040
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.TextBox txtPhone 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   2280
         TabIndex        =   14
         Top             =   2520
         Width           =   7335
      End
      Begin VB.TextBox txtLicense 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   2280
         TabIndex        =   5
         Top             =   3120
         Width           =   7335
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   2280
         TabIndex        =   4
         ToolTipText     =   "Full Name of the Driver"
         Top             =   720
         Width           =   7335
      End
      Begin VB.TextBox txtAddress 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   2280
         TabIndex        =   3
         ToolTipText     =   "Complete Address of the Driver"
         Top             =   1320
         Width           =   7335
      End
      Begin MSComCtl2.DTPicker dtBirth 
         Height          =   405
         Left            =   2280
         TabIndex        =   6
         Top             =   1920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   0
         CalendarForeColor=   -2147483637
         CalendarTitleBackColor=   0
         CalendarTitleForeColor=   -2147483637
         Format          =   105512961
         CurrentDate     =   43032
      End
      Begin MSComCtl2.DTPicker dtReg 
         Height          =   405
         Left            =   6600
         TabIndex        =   33
         Top             =   1920
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   0
         CalendarForeColor=   -2147483637
         CalendarTitleBackColor=   -2147483630
         CalendarTitleForeColor=   -2147483637
         Format          =   105512961
         CurrentDate     =   43023
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   360
         TabIndex        =   24
         Top             =   6120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Top             =   5640
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trip:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6360
         TabIndex        =   22
         Top             =   4920
         Width           =   465
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Plate No."
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   4920
         Width           =   1050
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "ASSIGNED VAN"
         BeginProperty Font 
            Name            =   "Century Schoolbook"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   0
         TabIndex        =   19
         Top             =   4320
         Width           =   9960
      End
      Begin VB.Label Label224 
         BackColor       =   &H00000000&
         Caption         =   "Civil Status:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   780
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   5160
         TabIndex        =   11
         Top             =   3360
         Width           =   990
      End
      Begin VB.Label Label900 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1230
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Reg:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   5280
         TabIndex        =   9
         Top             =   1920
         Width           =   1185
      End
      Begin VB.Label Label20 
         BackColor       =   &H00000000&
         Caption         =   "License:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No:"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Width           =   1185
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   720
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   149
      ImageHeight     =   57
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":5635
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":974B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":D08D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":10FE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":149CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":18F11
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":1C9AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":208F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":24243
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":27E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":2B584
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":2FEC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":33FB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrivers.frx":388C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvDriver 
      Height          =   5775
      Left            =   10200
      TabIndex        =   26
      Top             =   1800
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   12632256
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
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Driver's ID"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Address"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Birthdate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Phone"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "License"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Date Registered"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Civil Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "TIN"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "SSS"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Plate No."
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Type"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Make"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Trip"
         Object.Width           =   2187
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10320
      Top             =   1200
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
            Picture         =   "frmDrivers.frx":3C964
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   16800
      TabIndex        =   37
      Top             =   8160
      Width           =   585
   End
   Begin VB.Image imgPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   15720
      Picture         =   "frmDrivers.frx":3CEFE
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   15120
      TabIndex        =   34
      Top             =   1200
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   8040
      TabIndex        =   32
      Top             =   8175
      Width           =   1050
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   10920
      TabIndex        =   31
      Top             =   8175
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   13800
      TabIndex        =   30
      Top             =   8160
      Width           =   1005
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   5280
      TabIndex        =   29
      Top             =   8175
      Width           =   705
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   330
      Left            =   2520
      TabIndex        =   28
      Top             =   8175
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   1095
      Left            =   360
      Picture         =   "frmDrivers.frx":40FD8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Drivers Registration Management"
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
      TabIndex        =   1
      Top             =   300
      Width           =   7605
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   20655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   27
      Top             =   1080
      Width           =   7455
   End
   Begin VB.Image imgAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   1440
      Picture         =   "frmDrivers.frx":42C8F
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Image imgDel 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   9960
      Picture         =   "frmDrivers.frx":465C1
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Image imgUpdate 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   7080
      Picture         =   "frmDrivers.frx":4A04F
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Image imgSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   4200
      Picture         =   "frmDrivers.frx":4D7C3
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Image imgRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   12840
      Picture         =   "frmDrivers.frx":5119B
      Top             =   7920
      Width           =   2295
   End
End
Attribute VB_Name = "frmDrivers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim drivers_choice As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim DataArray(1 To 1000, 1 To 11) As Variant
Dim r As Integer
Dim NumberOfRows As Integer

Private Sub cmdPD_Click()
If drivers_choice = "" Then
            MsgBox "Please select a record to print.", vbCritical, "Warning"
        Else
        Set rs = New ADODB.Recordset
        rs.Open "Select * from Driver where DriverName='" & drivers_choice & "'", conn, 3, 3
            If rs.RecordCount > 0 Then
                With rptDrivers
                    Set rptDrivers.DataSource = rs
                        .Sections("Section1").Controls("Text1").DataField = "DriverID"
                        .Sections("Section1").Controls("Text2").DataField = "DriverName"
                        .Sections("Section1").Controls("Text5").DataField = "Address"
                        .Sections("Section1").Controls("Text3").DataField = "BDate"
                        .Sections("Section1").Controls("Text8").DataField = "PhoneNo"
                        .Sections("Section1").Controls("Text4").DataField = "License"
                        .Sections("Section1").Controls("Text6").DataField = "DateReg"
                        .Sections("Section1").Controls("Text9").DataField = "CStatus"
'                        .Sections("Section1").Controls("Text7").DataField = "TIN"
'                        .Sections("Section1").Controls("Text10").DataField = "SSS"
                        .Sections("Section1").Controls("Text12").DataField = "PlateNo"
                        .Show 1
                    Set rs = Nothing
                End With
            End If
        End If
    
End Sub

Private Sub Form_Load()
modConnect.Connected

Call Driver_Lock
Call PopBrand
Call PopDriveColor
Call RefDriver
Call PopCarInfo
Call imgRefresh_Click

cboCivStat.AddItem "Single"
cboCivStat.AddItem "Married"
cboCivStat.AddItem "Separated"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
End Sub
Private Sub imgAdd_Click()
   Call Driver_Unlock
   Call Driver_Clear
   imgUpdate.Enabled = False
   imgDel.Enabled = False
   imgSave.Enabled = True
End Sub

Private Sub imgAdd_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(1).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
End Sub

Private Sub imgDel_Click()
 If drivers_choice = vbNullString Then
            MsgBox "Please choose a record in the list.", vbExclamation, "Warning!"
        Else
         If MsgBox("Delete?", vbQuestion + vbYesNo) = vbYes Then

            Set rs = New ADODB.Recordset
            rs.Open "Select * from Driver where DriverName='" & drivers_choice & "'", conn, adOpenKeyset, adLockPessimistic
            With rs
                .Delete
                Call RefDriver
'                Call NoDriver
                .Close
            End With
            Set rs = Nothing
            MsgBox "Record Successfully Deleted!", vbInformation, "Success Deleted!"
            Call Driver_Clear
        End If
    End If
Call Driver_Lock
Call imgRefresh_Click
End Sub

Private Sub imgDel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(5).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
End Sub

Private Sub imgPrint_Click()
If Me.lvDriver.ListItems.Count = 0 Then Exit Sub
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Open(App.Path & "\Excel\Drivers.xlsx")


Set rs = New ADODB.Recordset
rs.Open "Select * from Driver", conn, adOpenDynamic, adLockOptimistic
NumberOfRows = rs.RecordCount
rs.MoveFirst
For r = 1 To NumberOfRows
DataArray(r, 1) = rs.Fields("DriverID")
DataArray(r, 2) = rs.Fields("DriverName")
DataArray(r, 3) = rs.Fields("Address")
DataArray(r, 4) = rs.Fields("PhoneNo")
DataArray(r, 5) = rs.Fields("License")
DataArray(r, 6) = rs.Fields("PlateNo")
rs.MoveNext
Next

Set oSheet = oBook.Worksheets(1)
oSheet.Range("E5").Resize(NumberOfRows, 6).Value = DataArray

oExcel.Visible = True
oExcel.Application.ActiveSheet.PrintPreview
rs.MoveFirst
Set rs = Nothing
End Sub

Private Sub imgPrint_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(11).Picture
End Sub

Private Sub imgRefresh_Click()
imgSave.Enabled = False
imgUpdate.Enabled = False
imgDel.Enabled = False
Call Driver_Clear
Call Driver_Lock
End Sub

Private Sub imgRefresh_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(7).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
End Sub

Private Sub imgSave_Click()
If txtName.Text = "" Or txtAddress.Text = "" Or txtPhone.Text = "" Or txtLicense.Text = "" Or cboCivStat.Text = "" Or cboPlate.Text = "" Then
            MsgBox "Some of your fields is empty. Please complete the information.", vbExclamation + vbOKOnly, "Warning"
              Call Driver_Unlock
        Exit Sub
        
        Else
            Set rs = New ADODB.Recordset
            rs.Open "Select * from Driver", conn, adOpenKeyset, adLockPessimistic
                With rs
                    Dim a As String
                    rs.MoveLast
                    a = Mid$(rs.Fields("DriverID"), 9, 13)
                    a = a + 1
                    .AddNew
                    .Fields("DriverID") = "D-ID" & Year(Now) & a
                    .Fields("DriverName") = txtName.Text
                    .Fields("Address") = txtAddress.Text
                    .Fields("BDate") = dtBirth.Value
                    .Fields("PhoneNo") = txtPhone.Text
                    .Fields("License") = txtLicense.Text
                    .Fields("DateReg") = dtReg.Value
                    .Fields("CStatus") = cboCivStat.Text
                    .Fields("TIN") = txtTIN.Text
                    .Fields("SSS") = txtSSS.Text
                    .Fields("PlateNo") = cboPlate.Text
                    .Fields("Type") = cboType.Text
                    .Fields("Make") = cboMake.Text
                    .Fields("Trip") = cboColor.Text
                    On Error Resume Next
                    .Update ' why man mag debug sya dri?
                End With
                Call RefDriver
'                Call NoDriver
            'rs.Close
            Set rs = Nothing
            MsgBox "Record Successfully Saved!", vbInformation, "Success Saved!"
            Call Driver_Clear
            
        End If
imgDel.Enabled = False
Call Driver_Lock
End Sub

Private Sub imgSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(3).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
End Sub

Private Sub imgUpdate_Click()
    If txtName.Text = "" Or txtAddress.Text = "" Or txtPhone.Text = "" Or txtLicense.Text = "" Or cboCivStat.Text = "" Or cboPlate.Text = "" Then
            MsgBox "Some of your fields is empty. Please complete the information.", vbExclamation + vbOKOnly, "Warning"
Call Driver_Unlock
        Exit Sub
        
        Else
            Set rs = New ADODB.Recordset
                rs.Open "Select * from Driver where DriverName='" & drivers_choice & "'", conn, adOpenDynamic, adLockOptimistic
            With rs
                !DriverName = txtName.Text
                !Address = txtAddress.Text
                !BDate = dtBirth.Value
                !PhoneNo = txtPhone.Text
                !License = txtLicense.Text
                !DateReg = dtReg.Value
                !CStatus = cboCivStat.Text
                !TIN = txtTIN.Text
                !SSS = txtSSS.Text
                !PlateNo = cboPlate.Text
                !Type = cboType.Text
                !Make = cboMake.Text
                !Trip = cboColor.Text
                rs.Update
                Call RefDriver
'                Call NoDriver
            End With
            Set rs = Nothing
            MsgBox "Record Successfully Updated!", vbInformation, "Success Updated!"
       
       End If
Call Driver_Clear
Call Driver_Lock
imgDel.Enabled = False
End Sub

Private Sub imgUpdate_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(9).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
End Sub

Private Sub lvDriver_Click()
imgSave.Enabled = False
imgUpdate.Enabled = True
imgDel.Enabled = True

Call Driver_Unlock
On Error Resume Next
drivers_choice = lvDriver.SelectedItem.SubItems(1)
txtName.Text = lvDriver.SelectedItem.SubItems(1)
txtAddress.Text = lvDriver.SelectedItem.SubItems(2)
dtBirth.Value = lvDriver.SelectedItem.SubItems(3)
txtPhone.Text = lvDriver.SelectedItem.SubItems(4)
txtLicense.Text = lvDriver.SelectedItem.SubItems(5)
dtReg.Value = lvDriver.SelectedItem.SubItems(6)
cboCivStat.Text = lvDriver.SelectedItem.SubItems(7)
txtTIN.Text = lvDriver.SelectedItem.SubItems(8)
txtSSS.Text = lvDriver.SelectedItem.SubItems(9)
cboPlate.Text = lvDriver.SelectedItem.SubItems(10)
cboType.Text = lvDriver.SelectedItem.SubItems(11)
cboMake.Text = lvDriver.SelectedItem.SubItems(12)
cboColor.Text = lvDriver.SelectedItem.SubItems(13)
End Sub

Private Sub Timer1_Timer()
If val(Me.lvDriver.ListItems.Count) > 1 Then
Me.Label15.Caption = "There are " & Me.lvDriver.ListItems.Count & " drivers found in the list."
ElseIf Me.lvDriver.ListItems.Count = 1 Then
Me.Label15.Caption = "There is " & Me.lvDriver.ListItems.Count & " drivers found in the list."
Else
Me.Label15.Caption = "No driver found in the list."
End If
End Sub
