VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vans"
   ClientHeight    =   9045
   ClientLeft      =   270
   ClientTop       =   720
   ClientWidth     =   17325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmCar.frx":0000
   ScaleHeight     =   9045
   ScaleWidth      =   17325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8280
      Top             =   480
   End
   Begin VB.TextBox txtSpeed 
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
      TabIndex        =   17
      Top             =   5400
      Width           =   6495
   End
   Begin VB.Frame FrameCar 
      BackColor       =   &H00000000&
      Caption         =   "Car Information"
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
      Height          =   5655
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   9135
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
         Left            =   7200
         TabIndex        =   26
         Top             =   2640
         Width           =   1695
      End
      Begin VB.ComboBox cboStatus 
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
         Left            =   2400
         TabIndex        =   19
         Top             =   5040
         Width           =   6495
      End
      Begin VB.ComboBox cboCondition 
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
         Left            =   2400
         TabIndex        =   18
         Top             =   4440
         Width           =   6495
      End
      Begin VB.ComboBox cboMake 
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
         Left            =   2400
         TabIndex        =   16
         Top             =   3240
         Width           =   6495
      End
      Begin VB.TextBox txtModel 
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
         Left            =   8880
         TabIndex        =   14
         Top             =   3120
         Visible         =   0   'False
         Width           =   6495
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
         Left            =   2400
         TabIndex        =   13
         Top             =   600
         Width           =   6495
      End
      Begin VB.TextBox txtReg 
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
         TabIndex        =   12
         Top             =   1320
         Width           =   6495
      End
      Begin VB.TextBox txtType 
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
         TabIndex        =   11
         Top             =   2040
         Width           =   6495
      End
      Begin MSComCtl2.DTPicker dtManu 
         Height          =   405
         Left            =   2760
         TabIndex        =   15
         Top             =   2640
         Width           =   2295
         _ExtentX        =   4048
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
         Format          =   107151361
         CurrentDate     =   41927
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sequence of Trip:"
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
         Left            =   5160
         TabIndex        =   25
         Top             =   2640
         Width           =   1995
      End
      Begin VB.Label Label99 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Plate No."
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
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reg No."
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
         Left            =   480
         TabIndex        =   9
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label Label101 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
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
         Left            =   600
         TabIndex        =   8
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Manufactured:"
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
         TabIndex        =   7
         Top             =   2640
         Width           =   2445
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "Model:"
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
         Left            =   600
         TabIndex        =   6
         Top             =   3240
         Width           =   795
      End
      Begin VB.Label Label155 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Speed:"
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
         Left            =   600
         TabIndex        =   5
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label Label16 
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
         Left            =   480
         TabIndex        =   4
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Availability Status:"
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
         TabIndex        =   3
         Top             =   5040
         Width           =   2145
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   360
      Top             =   7560
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
            Picture         =   "frmCar.frx":5601
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":9717
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":D059
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":10FB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":14999
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":18EDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":1C97B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":208BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":2420F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":27DCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":2B550
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":2FE95
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":33F7F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCar.frx":38894
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   360
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
            Picture         =   "frmCar.frx":3C930
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvCars 
      Height          =   5655
      Left            =   9480
      TabIndex        =   32
      Top             =   1920
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9975
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Card ID"
         Object.Width           =   19
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Plate Number"
         Object.Width           =   3422
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Registration No"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Trip"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Date Manufactured"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Model"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Model"
         Object.Width           =   3951
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Speed"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Condition"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Availability"
         Object.Width           =   4304
      EndProperty
   End
   Begin VB.Label Label3 
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
      Left            =   15840
      TabIndex        =   33
      Top             =   8160
      Width           =   585
   End
   Begin VB.Image imgPrint 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   14760
      Picture         =   "frmCar.frx":3CECA
      Top             =   7920
      Width           =   2295
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
      Left            =   13200
      TabIndex        =   0
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   9480
      TabIndex        =   31
      Top             =   1200
      Width           =   7575
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   600
      Picture         =   "frmCar.frx":40FA4
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   7650
      TabIndex        =   29
      Top             =   720
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Available Van(s)"
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
      Left            =   6975
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   6240
      TabIndex        =   27
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label25 
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
      Left            =   12960
      TabIndex        =   22
      Top             =   8160
      Width           =   1005
   End
   Begin VB.Label Label26 
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
      Left            =   10320
      TabIndex        =   23
      Top             =   8160
      Width           =   900
   End
   Begin VB.Label Label23 
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
      Left            =   2040
      TabIndex        =   20
      Top             =   8160
      Width           =   615
   End
   Begin VB.Image imgAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   960
      Picture         =   "frmCar.frx":4297F
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Image imgDel 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   9240
      Picture         =   "frmCar.frx":462B1
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Image imgRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   12000
      Picture         =   "frmCar.frx":49D3F
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label Label28 
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
      Left            =   7440
      TabIndex        =   24
      Top             =   8160
      Width           =   1050
   End
   Begin VB.Label Label24 
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
      Left            =   4800
      TabIndex        =   21
      Top             =   8160
      Width           =   705
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Van Records"
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
      Left            =   1800
      TabIndex        =   1
      Top             =   300
      Width           =   2790
   End
   Begin VB.Image imgUpdate 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   6480
      Picture         =   "frmCar.frx":4D995
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Image imgSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   3720
      Picture         =   "frmCar.frx":51109
      Top             =   7920
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00404040&
      Height          =   735
      Left            =   0
      TabIndex        =   30
      Top             =   240
      Width           =   20655
   End
End
Attribute VB_Name = "frmCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cars As String, Trip As String, makes As String
Dim a As Integer
Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim DataArray(1 To 1000, 1 To 11) As Variant
Dim r As Integer
Dim NumberOfRows As Integer

Private Sub cmdLogout_Click()
Unload Me
End Sub

Private Sub CarLock()
txtPlate.Enabled = False
txtReg.Enabled = False
txtType.Enabled = False
cboColor.Enabled = False
dtManu.Enabled = False
txtModel.Enabled = False
cboMake.Enabled = False
txtSpeed.Enabled = False
cboCondition.Enabled = False
cboStatus.Enabled = False
End Sub

Private Sub cboMake_Change()
Me.txtModel.Text = cboMake.Text
End Sub

Private Sub cboMake_Click()
Me.txtModel.Text = cboMake.Text
End Sub

Private Sub Form_Load()
modConnect.Connected
Call CarLock
Call RefCar
'Call NoCars
Call NoCarAvail
Call NoCarCondition
Call PopColor
Call PopMake
Call imgRefresh_Click

cboCondition.AddItem "Good"
cboCondition.AddItem "Rough"
cboCondition.AddItem "Damaged"

cboStatus.AddItem "Available"
cboStatus.AddItem "Unavailable"
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

Private Sub imgAdd_Click()
Me.imgDel.Enabled = False
Me.imgUpdate.Enabled = False
Me.imgSave.Enabled = True
Call Car_Unlock
Call Car_Clear
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
         
 If cars = vbNullString Then
            MsgBox "Please choose a record in the list.", vbExclamation, "Warning!"
        Else
        
            If MsgBox("Delete?", vbQuestion + vbYesNo) = vbYes Then
            
            Set rs = New ADODB.Recordset
            rs.Open "Select * from Car where PlateNo='" & cars & "'", conn, adOpenKeyset, adLockPessimistic
            With rs
                .Delete
                Call RefCar
'                Call NoCars
                Call NoCarAvail
                Call NoCarCondition
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
If Me.lvCars.ListItems.Count = 0 Then Exit Sub
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Open(App.Path & "\Excel\Cars.xlsx")

Set rs = New ADODB.Recordset
rs.Open "Select * from Car", conn, adOpenDynamic, adLockOptimistic
NumberOfRows = rs.RecordCount
rs.MoveFirst
For r = 1 To NumberOfRows
DataArray(r, 1) = rs.Fields("PlateNo")
DataArray(r, 2) = rs.Fields("RegNo")
DataArray(r, 3) = rs.Fields("Model")
DataArray(r, 4) = rs.Fields("Condition")
DataArray(r, 5) = rs.Fields("Trip")
DataArray(r, 6) = rs.Fields("AvailStatus")
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
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub imgRefresh_Click()
 imgSave.Enabled = False
 imgDel.Enabled = False
 imgUpdate.Enabled = False
 Call Car_Clear
 Call CarLock
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

 If txtPlate.Text = "" Or txtReg.Text = "" Or txtType.Text = "" Or cboColor.Text = "" Or txtModel.Text = "" Or cboMake.Text = "" Or txtSpeed.Text = "" Or cboCondition.Text = "" Or cboStatus.Text = "" Then
            MsgBox "Some of your fields is empty. Please complete the information.", vbExclamation + vbOKOnly, "Warning"
        Call Car_Unlock
Exit Sub
        Else
            Set rs = New ADODB.Recordset
            rs.Open "Select * from Car", conn, adOpenKeyset, adLockPessimistic
                With rs
                    Dim a As String
                      
                    rs.MoveLast
                    a = Mid$(rs.Fields("CarID"), 9, 13)
                    a = a + 1
                    .AddNew
                    .Fields("CarID") = "C-ID" & Year(Now) & a
                    .Fields("PlateNo") = txtPlate.Text
                    .Fields("RegNo") = txtReg.Text
                    .Fields("Type") = txtType.Text
                    .Fields("Trip") = cboColor.Text
                    .Fields("DateManu") = dtManu.Value
                    .Fields("Model") = txtModel.Text
                    .Fields("Make") = cboMake.Text
                    .Fields("Speed") = txtSpeed.Text
                    .Fields("Condition") = cboCondition.Text
                    .Fields("AvailStatus") = cboStatus.Text
                    On Error Resume Next
                    .Update
                End With
                Call RefCar
'                Call NoCars
                Call NoCarAvail
                Call NoCarCondition
            'rs.Close
            Set rs = Nothing
            MsgBox "Record Successfully Saved!", vbInformation, "Success Saved!"
            Call Car_Clear
        End If

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

If txtPlate.Text = "" Or txtReg.Text = "" Or txtType.Text = "" Or cboColor.Text = "" Or txtModel.Text = "" Or cboMake.Text = "" Or txtSpeed.Text = "" Or cboCondition.Text = "" Or cboStatus.Text = "" Then
            MsgBox "Some of your fields is empty. Please complete the information.", vbExclamation + vbOKOnly, "Warning"
         Call Car_Unlock
Exit Sub
        Else
            Set rs = New ADODB.Recordset
            rs.Open "Select * from Car where PlateNo='" & cars & "'", conn, adOpenKeyset, adLockPessimistic
            With rs
                !PlateNo = txtPlate.Text
                !RegNo = txtReg.Text
                !Type = txtType.Text
                !DateManu = dtManu.Value
                !Trip = cboColor.Text
                !Model = txtModel.Text
                !Make = cboMake.Text
                !Speed = txtSpeed.Text
                !Condition = cboCondition.Text
                !AvailStatus = cboStatus.Text
                rs.Update
                .Close
            End With
            Set rs = Nothing
            MsgBox "Record Successfully Deleted!", vbInformation, "Success Updated!"
                Call RefCar
'                Call NoCars
                Call NoCarAvail
                Call NoCarCondition
        End If
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

Private Sub jcbutton1_Click()

End Sub

Private Sub lvCars_Click()
'Me.txtModel.Text = cboMake.Text
imgUpdate.Enabled = True
imgDel.Enabled = True
imgSave.Enabled = False
Call Car_Unlock
On Error Resume Next
cars = lvCars.SelectedItem.SubItems(1)
txtPlate.Text = lvCars.SelectedItem.SubItems(1)
txtReg.Text = lvCars.SelectedItem.SubItems(2)
txtType.Text = lvCars.SelectedItem.SubItems(3)
cboColor.Text = lvCars.SelectedItem.SubItems(4)
dtManu.Value = lvCars.SelectedItem.SubItems(5)
txtModel.Text = lvCars.SelectedItem.SubItems(6)
cboMake.Text = lvCars.SelectedItem.SubItems(7)
txtSpeed.Text = lvCars.SelectedItem.SubItems(8)
cboCondition.Text = lvCars.SelectedItem.SubItems(9)
cboStatus.Text = lvCars.SelectedItem.SubItems(10)
End Sub

Private Sub Timer1_Timer()
If val(Me.lvCars.ListItems.Count) > 1 Then
Me.Label15.Caption = "There are " & Me.lvCars.ListItems.Count & " vans found in the list."
ElseIf Me.lvCars.ListItems.Count = 1 Then
Me.Label15.Caption = "There is " & Me.lvCars.ListItems.Count & " vans found in the list."
Else
Me.Label15.Caption = "No van found in the list."
End If
End Sub

