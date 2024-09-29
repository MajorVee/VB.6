VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EADED7D9-DE8D-4857-BE54-AE4203812ECD}#39.0#0"; "tssOfficeMenu1c.ocx"
Begin VB.Form ControlPanel 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Panel - Trial Version"
   ClientHeight    =   7245
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin tssOfficeMenu.OfficeMenu OfficeMenu1 
      Left            =   4440
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      MenuBitmapCount =   18
      MenuBitmap:1    =   "Control Panel.frx":0000
      Masking:1       =   16777215
      MenuKey:1       =   "#mnuLogOut"
      MenuBitmap:2    =   "Control Panel.frx":0552
      Masking:2       =   16777215
      MenuKey:2       =   "#mnuAccounts"
      MenuBitmap:3    =   "Control Panel.frx":0AA4
      Masking:3       =   16777215
      MenuKey:3       =   "#mnuBackup"
      MenuBitmap:4    =   "Control Panel.frx":0FF6
      Masking:4       =   13550269
      MenuKey:4       =   "#mnuExit"
      MenuBitmap:5    =   "Control Panel.frx":1548
      Masking:5       =   16777215
      MenuKey:5       =   "#MnuPO"
      MenuBitmap:6    =   "Control Panel.frx":1A9A
      Masking:6       =   13684944
      MenuKey:6       =   "#mnuSalesRPT"
      MenuBitmap:7    =   "Control Panel.frx":1FEC
      Masking:7       =   13684944
      MenuKey:7       =   "#mnuRooms"
      MenuBitmap:8    =   "Control Panel.frx":253E
      Masking:8       =   9408399
      MenuKey:8       =   "#mnurType"
      MenuBitmap:9    =   "Control Panel.frx":2A90
      Masking:9       =   9408399
      MenuKey:9       =   "#mnuCustomers"
      MenuBitmap:10   =   "Control Panel.frx":2FE2
      Masking:10      =   9408399
      MenuKey:10      =   "#mnuIn"
      MenuBitmap:11   =   "Control Panel.frx":3534
      Masking:11      =   9408399
      MenuKey:11      =   "#mnuOUt"
      MenuBitmap:12   =   "Control Panel.frx":3A86
      Masking:12      =   9408399
      MenuKey:12      =   "#mnuReserve"
      MenuBitmap:13   =   "Control Panel.frx":3FD8
      Masking:13      =   9408399
      MenuKey:13      =   "#mnuIncome"
      MenuBitmap:14   =   "Control Panel.frx":452A
      Masking:14      =   8553090
      MenuKey:14      =   "#mnuExpense"
      MenuBitmap:15   =   "Control Panel.frx":4A7C
      Masking:15      =   9408399
      MenuKey:15      =   "#mnuSettings"
      MenuBitmap:16   =   "Control Panel.frx":4FCE
      Masking:16      =   9408399
      MenuKey:16      =   "#mnuStat"
      MenuBitmap:17   =   "Control Panel.frx":5520
      Masking:17      =   13684944
      MenuKey:17      =   "#mnuConfirm"
      MenuBitmap:18   =   "Control Panel.frx":5A72
      Masking:18      =   8553090
      MenuKey:18      =   "#mnuUserLog"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ArrowColor      =   0
      BackgroundColor =   16383228
      BorderColor     =   8029834
      CheckFillColor  =   15263457
      ArrowColor      =   0
      BackgroundColor =   16383228
      BorderColor     =   8029834
      CheckFillColor  =   15263457
      CheckMarkColor  =   0
      DisabledIconTintColor=   10727860
      IconShadowColor =   11705744
      MenuBarBackgroundColor=   14215660
      MenuBarTextColor=   0
      SelectedArrowColor=   0
      SelectedCheckBorderColor=   12937777
      SelectedCheckFillColor=   14857624
      SelectedMenuBorderColor=   12937777
      SelectedMenuFillColor=   15651521
      SelectedMenuTextColor=   0
      SeparatorColor  =   12108485
      SideBarGradientColor1=   16383228
      SideBarGradientColor2=   14215660
      TopLevelGradientColor1=   16383228
      TopLevelGradientColor2=   14215660
      SideBarColor    =   0
   End
   Begin MSComctlLib.ImageList ilList 
      Left            =   8760
      Top             =   5880
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
            Picture         =   "Control Panel.frx":5FC4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   8040
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6022
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6080
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":60DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":613C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":619A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":61F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6256
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":62B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6312
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6370
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":63CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":642C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":648A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":64E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6546
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgForm 
      Left            =   9240
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":65A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6602
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6660
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":66BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":671C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":677A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListLeft 
      Left            =   7440
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":67D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6836
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6894
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":68F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6950
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":69AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6A0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6AC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6B26
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6B84
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6CFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6D5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6E74
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6ED2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6F30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgListLV 
      Left            =   6840
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":6FEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":704A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":70A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":7106
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":7164
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":71C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":7220
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":727E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":72DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":733A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":7398
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":73F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":7454
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":74B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":7510
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":756E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":75CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":762A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":7688
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":76E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":7744
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgQuickLaunch 
      Left            =   3120
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":77A2
            Key             =   "rmstat"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":807C
            Key             =   "user"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":8956
            Key             =   "reserve"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":9230
            Key             =   "creserve"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":9B0A
            Key             =   "cin"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":A3E4
            Key             =   "logoff"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":ACBE
            Key             =   "shutdown"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":B598
            Key             =   "room"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":BE72
            Key             =   "cout"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":C74C
            Key             =   "transaction"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":D026
            Key             =   "customers"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":D900
            Key             =   "roomtype"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Control Panel.frx":E1DA
            Key             =   "set"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView listQuickLaunch 
      CausesValidation=   0   'False
      Height          =   5655
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   9975
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgQuickLaunch"
      SmallIcons      =   "imgQuickLaunch"
      ForeColor       =   128
      BackColor       =   16777215
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "Control Panel.frx":EAB4
      NumItems        =   0
   End
   Begin VB.Frame xFrame6 
      BackColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3975
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Short-cut"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   45
         Width           =   900
      End
      Begin VB.Image Image8 
         Height          =   315
         Left            =   0
         Picture         =   "Control Panel.frx":EC16
         Stretch         =   -1  'True
         Top             =   0
         Width           =   5580
      End
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   2760
      TabIndex        =   7
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Top             =   7800
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   360
      TabIndex        =   5
      Top             =   7800
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   2160
      TabIndex        =   4
      Top             =   7800
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reservation System"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BB5900&
      Height          =   585
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   3930
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   0
      Picture         =   "Control Panel.frx":F09C
      Top             =   6960
      Width           =   12960
   End
   Begin VB.Image Image3 
      Height          =   645
      Left            =   0
      Picture         =   "Control Panel.frx":F69D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12960
   End
   Begin VB.Image Image2 
      Height          =   8850
      Left            =   3360
      Picture         =   "Control Panel.frx":10091
      Top             =   600
      Width           =   24000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLogOut 
         Caption         =   "Log-Off"
      End
      Begin VB.Menu mnuAccounts 
         Caption         =   "User Accounts"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup Data"
      End
      Begin VB.Menu mnuUserLog 
         Caption         =   "User Log"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuTrans 
      Caption         =   "Query"
      Begin VB.Menu mnuRooms 
         Caption         =   "Rooms"
      End
      Begin VB.Menu mnurType 
         Caption         =   "Room Type"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomers 
         Caption         =   "Customers"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIn 
         Caption         =   "Check-in"
      End
      Begin VB.Menu mnuOUt 
         Caption         =   "Check-out"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReserve 
         Caption         =   "Reservation"
      End
      Begin VB.Menu mnuConfirm 
         Caption         =   "Confirm Reservation"
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIncome 
         Caption         =   "Income Report"
      End
      Begin VB.Menu mnuExpense 
         Caption         =   "Expense Report"
      End
      Begin VB.Menu mnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStat 
         Caption         =   "Room Status"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Help"
      Begin VB.Menu MnuPO 
         Caption         =   "System Info"
      End
      Begin VB.Menu mnuSalesRPT 
         Caption         =   "Company Info"
      End
   End
End
Attribute VB_Name = "ControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AddQuickLaunchItems()
    
    listQuickLaunch.ListItems.Clear
    
    listQuickLaunch.ListItems.Add _
    , "RoomStat", "Room Status", "rmstat", "rmstat"

    listQuickLaunch.ListItems.Add _
    , "Reserve", "Reserve", "reserve", "reserve"

 
    listQuickLaunch.ListItems.Add _
    , "ConfirmReserve", "Confirm Reserve", "creserve", "creserve"
    
    listQuickLaunch.ListItems.Add _
    , "CheckIN", "Check-In", "cin", "cin"
    
     listQuickLaunch.ListItems.Add _
    , "CheckOut", "Check-Out", "cout", "cout"
    
    listQuickLaunch.ListItems.Add _
    , "Customers", "Customers", "user", "user"
    
    listQuickLaunch.ListItems.Add _
    , "Rooms", "Rooms", "room", "room"
    
    listQuickLaunch.ListItems.Add _
    , "RoomType", "Room Type", "roomtype", "roomtype"
    
    listQuickLaunch.ListItems.Add _
    , "Income", "Income", "transaction", "transaction"
    
   listQuickLaunch.ListItems.Add _
    , "Expense", "Expense", "customers", "customers"


    listQuickLaunch.ListItems.Add _
    , "HeadExceed", "Settings", "set", "set"
    
    ' listQuickLaunch.ListItems.Add _
    ', "PrintSchedule", "Class Schedule", "printschedule", "printschedule"
    
    
  ''''''''''''''''''''''''''''''''''''''''''''
    
    'listQuickLaunch.ListItems.Add _
   ' , "Logoff", "Log-off", "logoff", "logoff"
    
   ' listQuickLaunch.ListItems.Add _
  '  , "Shutdown", "Shutdown", "shutdown", "shutdown"

End Sub

Private Sub AddQuickLaunchItemsLimited()
    
    listQuickLaunch.ListItems.Clear
    
    listQuickLaunch.ListItems.Add _
    , "RoomStat", "Room Status", "rmstat", "rmstat"

    listQuickLaunch.ListItems.Add _
    , "Reserve", "Reserve", "reserve", "reserve"

 
    listQuickLaunch.ListItems.Add _
    , "ConfirmReserve", "Confirm Reserve", "creserve", "creserve"
    
    listQuickLaunch.ListItems.Add _
    , "CheckIN", "Check-In", "cin", "cin"
    
     listQuickLaunch.ListItems.Add _
    , "CheckOut", "Check-Out", "cout", "cout"
    
    listQuickLaunch.ListItems.Add _
    , "Customers", "Customers", "user", "user"
    
    listQuickLaunch.ListItems.Add _
    , "Rooms", "Rooms", "room", "room"
    
    listQuickLaunch.ListItems.Add _
    , "RoomType", "Room Type", "roomtype", "roomtype"
    
   ' listQuickLaunch.ListItems.Add _
    , "Income", "Income", "transaction", "transaction"
    
  ' listQuickLaunch.ListItems.Add _
    , "Expense", "Expense", "customers", "customers"


 '   listQuickLaunch.ListItems.Add _
  '  , "HeadExceed", "Settings", "set", "set"
    
    ' listQuickLaunch.ListItems.Add _
    ', "PrintSchedule", "Class Schedule", "printschedule", "printschedule"
    
    
  ''''''''''''''''''''''''''''''''''''''''''''
    
    'listQuickLaunch.ListItems.Add _
   ' , "Logoff", "Log-off", "logoff", "logoff"
    
   ' listQuickLaunch.ListItems.Add _
  '  , "Shutdown", "Shutdown", "shutdown", "shutdown"

End Sub
Private Sub Form_Activate()
If ControlPanel.Tag = "admin" Then
mnuAccounts.Enabled = True
mnuBackup.Enabled = True
mnuSettings.Enabled = True
mnuIncome.Enabled = True
mnuExpense.Enabled = True
AddQuickLaunchItems
Else
mnuAccounts.Enabled = False
mnuBackup.Enabled = False
mnuSettings.Enabled = False
mnuIncome.Enabled = False
mnuExpense.Enabled = False
AddQuickLaunchItemsLimited
End If
End Sub

Private Sub Form_Load()
If ControlPanel.Tag = "admin" Then
AddQuickLaunchItems
Else
AddQuickLaunchItemsLimited
End If
End Sub

Private Sub listQuickLaunch_DblClick()
Select Case listQuickLaunch.SelectedItem.Key
Case "RoomStat"
RoomStatusFrm.Show
Case "Reserve"
RoomReserveFrm.Show
Case "ConfirmReserve"
Case "CheckIN"
CheckInFrm.Show
Case "Customers"
CustomerFrm.Show
Case "Rooms"
RoomFrm.Show
Case "RoomType"
RoomTypeFrm.Show
Case "HeadExceed"
Case "Logoff"
If vbYes = MsgBox("Log-off?", vbQuestion + vbYesNo, "") Then
Unload Me
FrmLogin.Show
End If
Case "Shutdown"
If vbYes = MsgBox("Shutdown?", vbQuestion + vbYesNo, "") Then
End
End If
Case "CheckOut"
CheckOutFrm.Show
Case "Income"
IncomeFrm.Show
Case "Expense"
ExpenseFrm.Show
End Select
End Sub

Private Sub mnuAccounts_Click()
AccountsFrm.Show
End Sub





Private Sub mnuCustomers_Click()
CustomerFrm.Show
End Sub

Private Sub mnuExit_Click()
If vbYes = MsgBox("Shutdown?", vbQuestion + vbYesNo, "") Then
End
End If
End Sub

Private Sub mnuExpense_Click()
ExpenseFrm.Show
End Sub

Private Sub mnuIn_Click()
CheckInFrm.Show
End Sub

Private Sub mnuIncome_Click()
IncomeFrm.Show
End Sub

Private Sub mnuLogOut_Click()
If vbYes = MsgBox("Log-off?", vbQuestion + vbYesNo, "") Then
Unload Me
LoginFrm.Show
End If
End Sub

Private Sub mnuOUt_Click()
CheckOutFrm.Show
End Sub

Private Sub MnuPO_Click()
SysInfo.Show
End Sub

Private Sub mnuReserve_Click()
RoomReserveFrm.Show
End Sub
Private Sub mnuRooms_Click()
RoomFrm.Show
End Sub
Private Sub mnurType_Click()
RoomTypeFrm.Show
End Sub

Private Sub mnuSalesRPT_Click()
CompanyFrm.Show
End Sub



Private Sub mnuStat_Click()
RoomStatusFrm.Show
End Sub

Private Sub mnuUserLog_Click()
UserLogFrm.Show
End Sub
