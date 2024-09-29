VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Clients"
   ClientHeight    =   9435
   ClientLeft      =   270
   ClientTop       =   720
   ClientWidth     =   19275
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
   Picture         =   "frmClient.frx":0000
   ScaleHeight     =   9435
   ScaleWidth      =   19275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8520
      Top             =   960
   End
   Begin MSComCtl2.DTPicker dtBirth 
      Height          =   405
      Left            =   2520
      TabIndex        =   33
      Top             =   5160
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
      CalendarBackColor=   14737632
      Format          =   108134401
      CurrentDate     =   33170
   End
   Begin VB.ComboBox cboGender 
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
      Left            =   6240
      TabIndex        =   32
      Top             =   5160
      Width           =   2415
   End
   Begin VB.TextBox txtAddress 
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
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   31
      Top             =   3960
      Width           =   6135
   End
   Begin VB.TextBox txtLastname 
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
      IMEMode         =   3  'DISABLE
      Left            =   2520
      TabIndex        =   30
      Top             =   3360
      Width           =   6135
   End
   Begin VB.TextBox txtMidname 
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
      IMEMode         =   3  'DISABLE
      Left            =   2520
      TabIndex        =   29
      Top             =   2760
      Width           =   6135
   End
   Begin VB.TextBox txtFirstname 
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
      Left            =   2520
      TabIndex        =   28
      Top             =   2160
      Width           =   6135
   End
   Begin VB.TextBox txtOccup 
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
      Left            =   2520
      TabIndex        =   27
      Top             =   4560
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Client Information"
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
      Height          =   6735
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   8775
      Begin VB.TextBox txtPhone 
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
         IMEMode         =   3  'DISABLE
         Left            =   2400
         TabIndex        =   36
         Top             =   4320
         Width           =   6135
      End
      Begin VB.TextBox txtEmail 
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
         IMEMode         =   3  'DISABLE
         Left            =   2400
         TabIndex        =   35
         Top             =   4920
         Width           =   6135
      End
      Begin VB.TextBox txtLicense 
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
         Left            =   2400
         TabIndex        =   34
         Top             =   5520
         Width           =   6135
      End
      Begin MSComCtl2.DTPicker dtDateReg 
         Height          =   405
         Left            =   2400
         TabIndex        =   6
         Top             =   6120
         Width           =   6135
         _ExtentX        =   10821
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
         CalendarBackColor=   14737632
         Format          =   108134400
         CurrentDate     =   42979
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Driver's License:"
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
         Top             =   5520
         Width           =   1770
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
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
         Top             =   4920
         Width           =   675
      End
      Begin VB.Label Label1555 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No:"
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
         Top             =   4320
         Width           =   1185
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "Occupation:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Gender:"
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
         Left            =   5040
         TabIndex        =   16
         Top             =   3720
         Width           =   960
      End
      Begin VB.Label Label90 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Reg:"
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
         TabIndex        =   13
         Top             =   6120
         Width           =   1185
      End
      Begin VB.Label Label900 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth:"
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
         Top             =   3720
         Width           =   1515
      End
      Begin VB.Label Label13 
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
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
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
         TabIndex        =   10
         Top             =   2520
         Width           =   990
      End
      Begin VB.Label Label100 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
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
         TabIndex        =   9
         Top             =   1920
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Midlle Name:"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
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
         TabIndex        =   7
         Top             =   720
         Width           =   1275
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   240
      Top             =   8640
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
            Picture         =   "frmClient.frx":5601
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":9717
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":D059
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":10FB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":14999
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":18EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":1C97B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":208BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":2420F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":27DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":2B550
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":2FE95
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":33F7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClient.frx":38894
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvClient 
      Height          =   5895
      Left            =   9000
      TabIndex        =   25
      Top             =   2280
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   10398
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
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Client ID"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Firstname"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Midname"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Lastname"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Birth"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Gender"
         Object.Width           =   2363
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Occupation"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Phone"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Email"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "License"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Date Reg"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Postal"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Company"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Clearance"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label9 
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
      Left            =   15720
      TabIndex        =   39
      Top             =   8640
      Width           =   585
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Female Client(s)"
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
      Left            =   14265
      TabIndex        =   38
      Top             =   1695
      Width           =   1905
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   720
      Picture         =   "frmClient.frx":3C930
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Male Client(s)"
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
      Left            =   17400
      TabIndex        =   26
      Top             =   1695
      Width           =   1635
   End
   Begin VB.Image Image3 
      Height          =   855
      Left            =   13320
      Picture         =   "frmClient.frx":3DC4A
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   17985
      TabIndex        =   24
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Female Client(s)"
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
      Height          =   315
      Left            =   15000
      TabIndex        =   23
      Top             =   11700
      Width           =   1905
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   14985
      TabIndex        =   22
      Top             =   1320
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   16440
      Picture         =   "frmClient.frx":3F637
      Stretch         =   -1  'True
      Top             =   1230
      Width           =   840
   End
   Begin VB.Label Label15 
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
      Left            =   11040
      TabIndex        =   21
      Top             =   1560
      Width           =   285
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Client Registration Management"
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
      TabIndex        =   15
      Top             =   315
      Width           =   7305
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   0
      TabIndex        =   14
      Top             =   240
      Width           =   20655
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
      Left            =   1920
      TabIndex        =   1
      Top             =   8640
      Width           =   615
   End
   Begin VB.Label Label5 
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
      Left            =   4680
      TabIndex        =   2
      Top             =   8640
      Width           =   705
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
      Left            =   12840
      TabIndex        =   4
      Top             =   8640
      Width           =   1005
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
      Left            =   10080
      TabIndex        =   3
      Top             =   8640
      Width           =   900
   End
   Begin VB.Label Label18 
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
      Left            =   7320
      TabIndex        =   0
      Top             =   8640
      Width           =   1050
   End
   Begin VB.Image imgRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   11880
      Picture         =   "frmClient.frx":41185
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Image imgDel 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   9120
      Picture         =   "frmClient.frx":44DDB
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Image imgSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   3600
      Picture         =   "frmClient.frx":48869
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Image imgAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   840
      Picture         =   "frmClient.frx":4C241
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Image imgUpdate 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   6360
      Picture         =   "frmClient.frx":4FB73
      Top             =   8400
      Width           =   2295
   End
   Begin VB.Label Label10 
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
      Height          =   1095
      Left            =   9000
      TabIndex        =   37
      Top             =   1080
      Width           =   10215
   End
   Begin VB.Image imgPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   14640
      Picture         =   "frmClient.frx":532E7
      Top             =   8400
      Width           =   2295
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim client_mode As String
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim DataArray(1 To 1000, 1 To 11) As Variant
Dim r As Integer
Dim NumberOfRows As Integer


Private Sub cmdLogout_Click()
Unload Me
End Sub

Private Sub Form_Load()
modConnect.Connected
'
Call Client_Lock
Call RefClient
Call NoMaleClient
Call NoFemaleClient

'Call NoClient
cboGender.AddItem "Male"
cboGender.AddItem "Female"
Call imgRefresh_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub imgAdd_Click()
imgSave.Enabled = True
imgUpdate.Enabled = False
imgDel.Enabled = False
    Call Client_Clear
    Call Client_Unlock
End Sub

Private Sub imgAdd_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(1).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub imgClose_Click()
Unload Me
frmMain.Show
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(13).Picture
End Sub

Private Sub imgDel_Click()
If client_mode = vbNullString Then
            MsgBox "Please choose a record in the list.", vbExclamation, "Warning!"
        Else
        
        If MsgBox("Delete?", vbQuestion + vbYesNo) = vbYes Then
    
            Set rs = New ADODB.Recordset
            rs.Open "Select * from [Client] where FName='" & client_mode & "'", conn, adOpenKeyset, adLockPessimistic
            With rs
                .Delete
                Call RefClient
'                Call NoClient
                Call NoMaleClient
                Call NoFemaleClient
                .Close
            End With
            Set rs = Nothing
            MsgBox "Record Successfully Deleted!", vbInformation, "Success Deleted!"
        End If
    End If
Call imgRefresh_Click
End Sub

Private Sub imgDel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(5).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub imgPrint_Click()
If Me.lvClient.ListItems.Count = 0 Then Exit Sub
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Open(App.Path & "\Excel\Clients.xlsx")

Set rs = New ADODB.Recordset
rs.Open "Select * from Client", conn, adOpenDynamic, adLockOptimistic
NumberOfRows = rs.RecordCount
rs.MoveFirst
For r = 1 To NumberOfRows
DataArray(r, 1) = rs.Fields("FName")
DataArray(r, 2) = rs.Fields("LName")
DataArray(r, 3) = rs.Fields("Address")
DataArray(r, 4) = rs.Fields("Gender")
DataArray(r, 5) = rs.Fields("PhoneNo")
DataArray(r, 6) = rs.Fields("Email")
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

    Call Client_Clear
    Call Client_Lock
End Sub

Private Sub imgRefresh_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(7).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub imgSave_Click()
     If txtFirstname.Text = "" Or txtMidname.Text = "" Or txtAddress.Text = "" Or cboGender.Text = "" Or txtOccup.Text = "" Or txtPhone.Text = "" Or txtEmail.Text = "" Or txtLicense.Text = "" Then
            MsgBox "Some of your fields is empty. Please complete the information.", vbExclamation + vbOKOnly, "Warning"
        
        Call Client_Unlock
        Exit Sub
        Else
            Set rs = New ADODB.Recordset
            rs.Open "Select * from Client", conn, adOpenKeyset, adLockPessimistic
                With rs
                    Dim a As String
                    rs.MoveLast
                    
                    a = rs.Fields("ClientID")
                    a = a + 1
                    .AddNew
                    .Fields("ClientID") = "0000" & a
                    .Fields("FName") = txtFirstname.Text
                    .Fields("MName") = txtMidname.Text
                    .Fields("LName") = txtLastname.Text
                    .Fields("Address") = txtAddress.Text
                    .Fields("DOB") = dtBirth.Value
                    .Fields("Gender") = cboGender.Text
                    .Fields("Occup") = txtOccup.Text
                    .Fields("PhoneNo") = txtPhone.Text
                    .Fields("Email") = txtEmail.Text
                    .Fields("License") = txtLicense.Text
                    .Fields("DateReg") = dtDateReg.Value
                    .Update
                End With
                Call RefClient
'                Call NoClient
                Call NoMaleClient
                Call NoFemaleClient
            'rs.Close
            Set rs = Nothing
            MsgBox "Record Successfully Saved!", vbInformation, "Success Saved!"
            Call Client_Clear
        End If
        Call Client_Lock
End Sub

Private Sub imgSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(3).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(10).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub imgUpdate_Click()
 If txtFirstname.Text = "" Or txtMidname.Text = "" Or txtAddress.Text = "" Or cboGender.Text = "" Or txtOccup.Text = "" Or txtPhone.Text = "" Or txtEmail.Text = "" Or txtLicense.Text = "" Then
            MsgBox "Some of your fields is empty. Please complete the information.", vbExclamation + vbOKOnly, "Warning"
        Call Client_Unlock
        Exit Sub
        Else
            Set rs = New ADODB.Recordset
                rs.Open "Select * from [Client] where FName='" & client_mode & "'", conn, adOpenDynamic, adLockOptimistic
            With rs
                !FName = txtFirstname.Text
                !MName = txtMidname.Text
                !LName = txtLastname.Text
                !Address = txtAddress.Text
                !DOB = dtBirth.Value
                !Gender = cboGender.Text
                !Occup = txtOccup.Text
                !PhoneNo = txtPhone.Text
                !Email = txtEmail.Text
                !License = txtLicense.Text
                rs.Update
                Call RefClient
'                Call NoClient
                Call NoMaleClient
                Call NoFemaleClient
            End With
            Set rs = Nothing
            MsgBox "Record Successfully Updated!", vbInformation, "Success Updated!"
        End If
Call imgRefresh_Click
End Sub

Private Sub imgUpdate_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgUpdate.Picture = Me.ImageList2.ListImages(9).Picture
Me.imgPrint.Picture = Me.ImageList2.ListImages(12).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub lvClient_Click()
imgSave.Enabled = False
Me.imgDel.Enabled = True
Me.imgUpdate.Enabled = True
Call Client_Unlock
On Error Resume Next
client_mode = lvClient.SelectedItem.SubItems(1)
txtFirstname.Text = lvClient.SelectedItem.SubItems(1)
txtMidname.Text = lvClient.SelectedItem.SubItems(2)
txtLastname.Text = lvClient.SelectedItem.SubItems(3)
txtAddress.Text = lvClient.SelectedItem.SubItems(4)
dtBirth.Value = lvClient.SelectedItem.SubItems(5)
cboGender.Text = lvClient.SelectedItem.SubItems(6)
txtOccup.Text = lvClient.SelectedItem.SubItems(7)
txtPhone.Text = lvClient.SelectedItem.SubItems(8)
txtEmail.Text = lvClient.SelectedItem.SubItems(9)
txtLicense.Text = lvClient.SelectedItem.SubItems(10)
dtDateReg.Value = lvClient.SelectedItem.SubItems(11)
End Sub

Private Sub Timer1_Timer()
If val(Me.lvClient.ListItems.Count) > 1 Then
Me.Label15.Caption = "There are " & Me.lvClient.ListItems.Count & " clients found in the list."
ElseIf Me.lvClient.ListItems.Count = 1 Then
Me.Label15.Caption = "There is " & Me.lvClient.ListItems.Count & " clients found in the list."
Else
Me.Label15.Caption = "No client found in the list."
End If
End Sub
