VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRent 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Van Rental Processing Order"
   ClientHeight    =   9660
   ClientLeft      =   1050
   ClientTop       =   720
   ClientWidth     =   16650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frmRent.frx":0000
   ScaleHeight     =   9660
   ScaleWidth      =   16650
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8760
      Top             =   3600
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000012&
      Caption         =   "Cost Rent Summary"
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
      Height          =   3255
      Left            =   480
      TabIndex        =   24
      Top             =   4680
      Width           =   7815
      Begin VB.TextBox txtRateApplied 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1800
         TabIndex        =   30
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox txtTaxRate 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   1800
         TabIndex        =   29
         Text            =   "0.12"
         Top             =   1080
         Width           =   5775
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         TabIndex        =   35
         Top             =   2640
         Width           =   645
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   33
         Top             =   2550
         Width           =   5775
      End
      Begin VB.Label lblSubTotal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   32
         Top             =   2040
         Width           =   5775
      End
      Begin VB.Label lblTaxAmount 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1800
         TabIndex        =   31
         Top             =   1590
         Width           =   5775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Applied:"
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
         TabIndex        =   28
         Top             =   600
         Width           =   1620
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Rate:"
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
         TabIndex        =   27
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tax Amount:"
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
         TabIndex        =   26
         Top             =   1560
         Width           =   1500
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sub Total:"
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
         TabIndex        =   25
         Top             =   2160
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "Van Details:"
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
      Height          =   2415
      Left            =   8640
      TabIndex        =   14
      Top             =   960
      Width           =   7935
      Begin VB.ComboBox cboPlate 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   1680
         TabIndex        =   46
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox txtModel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   1200
         Width           =   6015
      End
      Begin VB.TextBox txtCondition 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1800
         Width           =   6015
      End
      Begin VB.TextBox txtMake 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   6375
      End
      Begin VB.TextBox txtMileage 
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
         TabIndex        =   20
         Top             =   3000
         Visible         =   0   'False
         Width           =   6375
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
         TabIndex        =   19
         Top             =   3600
         Visible         =   0   'False
         Width           =   6375
      End
      Begin VB.Label Label15 
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
         TabIndex        =   47
         Top             =   600
         Width           =   1050
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
         TabIndex        =   18
         Top             =   1200
         Width           =   1290
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
         TabIndex        =   17
         Top             =   1800
         Width           =   1215
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
         TabIndex        =   16
         Top             =   3000
         Visible         =   0   'False
         Width           =   1005
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
         TabIndex        =   15
         Top             =   3600
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Rental Order Processed For"
      BeginProperty Font 
         Name            =   "Century Schoolbook"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   8415
      Begin VB.TextBox txtDay 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2400
         Width           =   375
      End
      Begin VB.ComboBox cboClient 
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
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   600
         Width           =   6015
      End
      Begin VB.TextBox txtAddress 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1200
         Width           =   6015
      End
      Begin VB.ComboBox cboPass 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmRent.frx":5601
         Left            =   2160
         List            =   "frmRent.frx":5603
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3000
         Width           =   6015
      End
      Begin VB.ComboBox cboDriver 
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
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1800
         Width           =   6015
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   405
         Left            =   2160
         TabIndex        =   6
         Top             =   2400
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Helvetica"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   8421504
         CalendarForeColor=   -2147483637
         CalendarTitleBackColor=   12632256
         CalendarTrailingForeColor=   16761024
         Format          =   53215233
         CurrentDate     =   42736
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   405
         Left            =   4680
         TabIndex        =   7
         Top             =   2400
         Width           =   1935
         _ExtentX        =   3413
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
         CalendarBackColor=   8421504
         CalendarForeColor=   -2147483637
         CalendarTitleBackColor=   12632256
         CalendarTrailingForeColor=   16761024
         Format          =   53215233
         CurrentDate     =   42736
         MinDate         =   42005
      End
      Begin VB.Label Label27 
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
         Left            =   1320
         TabIndex        =   45
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client Address:"
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
         Left            =   140
         TabIndex        =   41
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   1530
      End
      Begin VB.Label Label156 
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
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   1650
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Trip to:"
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
         TabIndex        =   10
         Top             =   3000
         Width           =   780
      End
      Begin VB.Label Label29 
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
         Left            =   4200
         TabIndex        =   9
         Top             =   2400
         Width           =   330
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
         Left            =   6720
         TabIndex        =   8
         Top             =   2400
         Width           =   855
      End
   End
   Begin MSComctlLib.ListView lvTrans 
      Height          =   3855
      Left            =   8640
      TabIndex        =   34
      Top             =   4200
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6800
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Trans. No"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Receipt No"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Rate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Tax Rate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Tax Amount"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Sub Total"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   2716
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8160
      Top             =   6720
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
            Picture         =   "frmRent.frx":5605
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   120
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   149
      ImageHeight     =   57
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":5B9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":9CB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":D5F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":1154F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":14F37
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":1947B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":1CF19
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":20E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":247AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":2836A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":2BAEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":30433
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":3451D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":38E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":3CECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRent.frx":41326
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   12690
      TabIndex        =   42
      Top             =   3600
      Width           =   345
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   8640
      TabIndex        =   43
      Top             =   3480
      Width           =   7935
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay"
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
      Left            =   8400
      TabIndex        =   40
      Top             =   8760
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Image imgPay 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   7200
      Picture         =   "frmRent.frx":44F44
      Top             =   8520
      Visible         =   0   'False
      Width           =   2295
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
      Left            =   11040
      TabIndex        =   37
      Top             =   8760
      Width           =   1005
   End
   Begin VB.Image imgRefresh 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   10080
      Picture         =   "frmRent.frx":48B52
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Rent"
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
      Left            =   2280
      TabIndex        =   39
      Top             =   8760
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label Label10 
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
      Left            =   5400
      TabIndex        =   38
      Top             =   8760
      Width           =   705
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
      Left            =   14040
      TabIndex        =   36
      Top             =   8760
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   855
      Left            =   720
      Picture         =   "frmRent.frx":4C7A8
      Stretch         =   -1  'True
      Top             =   45
      Width           =   840
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Van Rent"
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
      Left            =   1680
      TabIndex        =   0
      Top             =   165
      Width           =   1830
   End
   Begin VB.Image imgDel 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   12960
      Picture         =   "frmRent.frx":4CE91
      Top             =   8520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image imgSave 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   4320
      Picture         =   "frmRent.frx":5091F
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Image imgAdd 
      BorderStyle     =   1  'Fixed Single
      Height          =   915
      Left            =   1440
      Picture         =   "frmRent.frx":542F7
      Top             =   8520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   0
      TabIndex        =   44
      Top             =   120
      Width           =   20655
   End
End
Attribute VB_Name = "frmRent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
Me.imgPay.Picture = Me.ImageList2.ListImages(16).Picture
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgPay.Picture = Me.ImageList2.ListImages(16).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub imgAdd_Click()
 Call Rent_Unlock
End Sub

Private Sub imgAdd_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(1).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgPay.Picture = Me.ImageList2.ListImages(16).Picture
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
'Me.imgClose.Picture = Me.ImageList2.ListImages(13).Picture
Me.imgPay.Picture = Me.ImageList2.ListImages(16).Picture
End Sub

Private Sub imgDel_Click()
   If Trans = vbNullString Then
            MsgBox "Please choose a record in the list to delete the transaction.", vbExclamation, "Warning!"
        Else
            Set rs = New ADODB.Recordset
            rs.Open "Select * from Order_Transaction where ReceiptNo='" & Trans & "'", conn, adOpenKeyset, adLockPessimistic
            With rs
                .Delete
                Call RefTrans
                .Close
            End With
            Set rs = Nothing
            MsgBox "Transaction Successfully Deleted!", vbInformation
        End If
    
End Sub

Private Sub imgDel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(5).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgPay.Picture = Me.ImageList2.ListImages(16).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub imgPay_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgPay.Picture = Me.ImageList2.ListImages(15).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub imgRefresh_Click()

Call Rent_Clear
Call PopClient
Call PopDriverName
Call PopCarPlate
Call RefTrans
Call PassLoad
End Sub

Private Sub imgRefresh_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(7).Picture
Me.imgPay.Picture = Me.ImageList2.ListImages(16).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub imgSave_Click()
  If cboClient.Text = "" Or txtAddress.Text = "" Or cboDriver.Text = "" Or cboPlate.Text = "" _
        Or txtModel.Text = "" Or txtCondition.Text = "" Or txtMake.Text = "" Or txtRateApplied.Text = "" Or txtTaxRate.Text = "" Or txtDay.Text = "" Or cboPass.Text = "" Then
            MsgBox "Some of your fields is empty. Please complete the information.", vbExclamation + vbOKOnly, "Warning"
        Else
        
        If Not IsNumeric(txtRateApplied.Text) < 0 Then
                MsgBox "Cannot be processed!", vbCritical, "Invalid Payment"
    
    Else
    
        Set rs = New ADODB.Recordset
            rs.Open "Select * from Purchase_Order", conn, adOpenKeyset, adLockPessimistic
                With rs
                    .AddNew
                    .Fields("ClientID") = client
                    .Fields("DriverID") = driver
                    .Fields("ClientName") = cboClient.Text
                    .Fields("DriverName") = cboDriver.Text
                    .Fields("DateFrom") = dtFrom.Value
                    .Fields("DateTo") = dtTo.Value
                    .Fields("NoDay") = txtDay.Text
                    .Fields("TripTo") = cboPass.Text
                    .Fields("ProcessedBy") = Label10.Caption
                    .Update
                End With
                Set rs = Nothing
                
            Set rs2 = New ADODB.Recordset
            rs2.Open "Select * from Rent_Car", conn, adOpenKeyset, adLockPessimistic
                With rs2
                    .AddNew
                    .Fields("ClientID") = client
                    .Fields("CarID") = car 'Label4.Caption
                    .Fields("DriverID") = driver
                    .Fields("PlateNo") = cboPlate.Text
                    .Fields("Model") = txtModel.Text
                    .Fields("Condition") = txtCondition.Text
                    .Fields("Make") = txtMake.Text
                    .Fields("MileAge") = txtMileage.Text
                    .Fields("TankLevel") = txtTank.Text
                    .Fields("DateRented") = Date
                    .Fields("Status") = "On Field"
                    On Error Resume Next
                    .Update
                End With
                Set rs2 = Nothing
                
            Set rs3 = New ADODB.Recordset
            rs3.Open "Select * from Order_Transaction", conn, adOpenKeyset, adLockPessimistic
                With rs3
                    Dim a As String, B As String
                    rs3.MoveLast
                    a = Mid$(rs3.Fields("ReceiptNo"), 9, 13)
                    a = a + 1
                    B = Mid$(rs3.Fields("TransNo"), 13, 18)
                    B = B + 1
                    
                    .AddNew
                    .Fields("CarID") = car 'Label4.Caption
                    .Fields("ReceiptNo") = "OR-1" & Year(Now) & a
                    .Fields("RateApplied") = txtRateApplied.Text
                    .Fields("TaxRate") = CCur(txtTaxRate.Text)
                    .Fields("TaxAmount") = lblTaxAmount.Caption
                    .Fields("SubTotal") = lblSubTotal.Caption
                    .Fields("Total") = lblTotal.Caption
                    .Fields("TransNo") = "TRANS-00" & Year(Now) & B
                    .Update
                End With
                Set rs3 = Nothing
            'Dim stat As TextBox
            'stat.Text = "Unavailable"
            Set rs4 = New ADODB.Recordset
            'rs4.Open "Select Car where PlateNo where ='" & cboPlate.Text & "'", conn, 3, 3
            
'            rs4.Open "Select * from Car where PlateNo='" & car & "'", conn, adOpenKeyset, adLockPessimistic
            rs4.Open "Select AvailStatus, PlateNo from Car where PlateNo='" & Me.cboPlate.Text & "'", conn, 3, 3
                With rs4
                      !AvailStatus = "Unavailable"
                    .Update
                End With
            Set rs4 = Nothing
            'rs.Close
            MsgBox "Record Successfully Saved!", vbInformation, "Success Saved!"
           
            Call RefTrans
        End If
End If
'Call imgRefresh_Click
'call
Set rs = New ADODB.Recordset
        rs.Open "Select * from Order_Transaction", conn, 3, 3
            'If rs.RecordCount > 0 Then
               With rptTrans
                    Set rptTrans.DataSource = rs
                        .Sections("Section1").Controls("Label2").Caption = client
                        .Sections("Section1").Controls("Label20").Caption = cboClient.Text
                        .Sections("Section1").Controls("label22").Caption = txtAddress.Text
                        .Sections("Section1").Controls("label21").Caption = frmMain.Label10.Caption
                        .Sections("Section1").Controls("Label23").Caption = txtDay.Text
                        .Sections("Section1").Controls("Label24").Caption = cboPass.Text
                        .Sections("Section1").Controls("Label26").Caption = driver
                        .Sections("Section1").Controls("Label28").Caption = cboDriver.Text
                        .Sections("Section1").Controls("Label29").Caption = cboPlate.Text
                        .Sections("Section1").Controls("Label30").Caption = txtModel.Text
                        .Sections("Section1").Controls("Label31").Caption = txtCondition.Text
'                        .Sections("Section1").Controls("Label32").Caption = txtMake.Text
'                        .Sections("Section1").Controls("Label33").Caption = txtMileage.Text
'                        .Sections("Section1").Controls("Label34").Caption = txtTank.Text
                        .Sections("Section1").Controls("Label40").Caption = txtRateApplied.Text
                        .Sections("Section1").Controls("Label41").Caption = txtTaxRate.Text
                        .Sections("Section1").Controls("Label42").Caption = lblTaxAmount.Caption
                        .Sections("Section1").Controls("Label43").Caption = lblSubTotal.Caption
                        .Sections("Section1").Controls("Label44").Caption = lblTotal.Caption
'                        .Sections("Section1").Controls("Label46").Caption = Format(txtCash.Text, "#,###,##0.00")
'                        .Sections("Section1").Controls("Label48").Caption = txtChange.Text
                        '.Sections("Section1").Controls("Label34").Caption = txtTank.Text
                        .Show 1
                    Set rs = Nothing
                End With
End Sub

Private Sub imgSave_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(3).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
Me.imgPay.Picture = Me.ImageList2.ListImages(16).Picture
End Sub


Private Sub lvTrans_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Me.imgAdd.Picture = Me.ImageList2.ListImages(2).Picture
Me.imgSave.Picture = Me.ImageList2.ListImages(4).Picture
Me.imgDel.Picture = Me.ImageList2.ListImages(6).Picture
Me.imgRefresh.Picture = Me.ImageList2.ListImages(8).Picture
Me.imgPay.Picture = Me.ImageList2.ListImages(16).Picture
'Me.imgClose.Picture = Me.ImageList2.ListImages(14).Picture
End Sub

Private Sub cboClient_Click()
Call Display_Rec
'txtAddress.SetFocus
End Sub

Private Sub cboDriver_Click()
Call Display_Plate
End Sub

'Private Sub cboPass_Click()
'txtMileage.SetFocus
'End Sub

Private Sub cboPlate_Change()
Call Display_Model
End Sub
Private Sub dtTo_Change()
txtDay.Text = DateDiff("d", dtFrom.Value, dtTo.Value)
End Sub

Private Sub Form_Load()
modConnect.Connected
Call PopClient
Call PopDriverName
Call PopCarPlate
Call RefTrans
Call PassLoad
'Call Rent_Lock
''Frame5.Visible = False
'Label10.Caption = MainForm.Label10.Caption


dtFrom.Value = Date
dtTo.Value = Date
End Sub
Private Sub PassLoad()
cboPass.AddItem "Ipil to Zamboanga"
cboPass.AddItem "Ipil to Cagayan"
cboPass.AddItem "Ipil to Dapitan"
cboPass.AddItem "Ipil to Pagadian"
cboPass.AddItem "Ipil to Dipolog"
End Sub


Private Sub lvTrans_Click()
On Error Resume Next
Trans = lvTrans.SelectedItem.SubItems(1)
End Sub

Private Sub txtCash_Change()
On Error Resume Next
txtChange.Text = Format((txtCash.Text) - (lblTotal.Caption), "###,###0.00")
End Sub

Private Sub txtCash_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    Call imgPay_Click
'End If
End Sub

Private Sub Timer1_Timer()
If val(Me.lvTrans.ListItems.Count) > 1 Then
Me.Label26.Caption = "There are " & Me.lvTrans.ListItems.Count & " transactions found in the list."
ElseIf Me.lvTrans.ListItems.Count = 1 Then
Me.Label26.Caption = "There is " & Me.lvTrans.ListItems.Count & " transactions found in the list."
Else
Me.Label26.Caption = "No transaction found in the list."
End If
End Sub

Private Sub txtRateApplied_Change()
On Error Resume Next
Dim a, B As Integer
Dim tax As Double

lblSubTotal.Caption = val(txtDay.Text) * val(txtRateApplied.Text)
lblTaxAmount.Caption = val(txtRateApplied.Text) * val(txtTaxRate.Text)

a = lblSubTotal.Caption
B = lblTaxAmount.Caption
lblTotal.Caption = Format((a + B), "###,###0.00")
End Sub

'Private Sub txtRateApplied_LostFocus()
'txtRateApplied.Text = Format(txtRateApplied.Text, "###,###0.00")
'End Sub

