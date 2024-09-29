VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGame 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Memory Game"
   ClientHeight    =   10215
   ClientLeft      =   -15
   ClientTop       =   630
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   10215
   ScaleWidth      =   14430
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "Rockwell"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      TabIndex        =   9
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   14520
      Top             =   1200
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   18120
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   18120
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   18120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6375
      Left            =   14520
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   11245
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   15960
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   125
      ImageHeight     =   182
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   30
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9ADD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":EA09
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16ED8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E143
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":253C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29767
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":301A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":372AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3BA18
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4062C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":48287
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4D152
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":51D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":565CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B3AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":61F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":66E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6F30D
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":76578
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":7D7FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":81B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":885D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8F6E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":93E4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":98A61
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A06BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A5587
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AA1BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AEA03
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B37E1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   15360
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   126
      ImageHeight     =   182
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BA347
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   6375
      Left            =   16200
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   11245
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ImageIndex"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   915
      Left            =   8520
      Picture         =   "Form1.frx":C6823
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2085
   End
   Begin VB.Image Image6 
      Height          =   915
      Left            =   3960
      Picture         =   "Form1.frx":C9F8D
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3045
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   960
      Width           =   1230
   End
   Begin VB.Image Image5 
      Height          =   975
      Left            =   360
      Picture         =   "Form1.frx":CD695
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3135
   End
   Begin VB.Image imgTimer 
      Height          =   1335
      Left            =   11760
      Picture         =   "Form1.frx":D0692
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Image Image4 
      Height          =   915
      Left            =   11760
      Picture         =   "Form1.frx":D2110
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2445
   End
   Begin VB.Label lblPoints 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7320
      TabIndex        =   8
      Top             =   240
      Width           =   405
   End
   Begin VB.Label lblTries 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   870
      Left            =   10680
      TabIndex        =   5
      Top             =   240
      Width           =   810
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   17
      Left            =   10320
      Picture         =   "Form1.frx":D6CA1
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   17
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   16
      Left            =   9000
      Picture         =   "Form1.frx":E316D
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   15
      Left            =   7680
      Picture         =   "Form1.frx":EF639
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   14
      Left            =   6360
      Picture         =   "Form1.frx":FBB05
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   13
      Left            =   5040
      Picture         =   "Form1.frx":107FD1
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   12
      Left            =   3720
      Picture         =   "Form1.frx":11449D
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   8
      Left            =   6360
      Picture         =   "Form1.frx":120969
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   1
      Left            =   5040
      Picture         =   "Form1.frx":12CE35
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   6
      Left            =   3720
      Picture         =   "Form1.frx":139301
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   11
      Left            =   10320
      Picture         =   "Form1.frx":1457CD
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   10
      Left            =   9000
      Picture         =   "Form1.frx":151C99
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   9
      Left            =   7680
      Picture         =   "Form1.frx":15E165
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   0
      Left            =   3720
      Picture         =   "Form1.frx":16A631
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   5
      Left            =   10320
      Picture         =   "Form1.frx":176AFD
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   3
      Left            =   7680
      Picture         =   "Form1.frx":182FC9
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   2
      Left            =   6360
      Picture         =   "Form1.frx":18F495
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   4
      Left            =   9000
      Picture         =   "Form1.frx":19B961
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   7
      Left            =   5040
      Picture         =   "Form1.frx":1A7E2D
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   555
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   2130
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   29
      Left            =   10320
      Picture         =   "Form1.frx":1B42F9
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   28
      Left            =   9000
      Picture         =   "Form1.frx":1C07C5
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   27
      Left            =   7680
      Picture         =   "Form1.frx":1CCC91
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   26
      Left            =   6360
      Picture         =   "Form1.frx":1D915D
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   25
      Left            =   5040
      Picture         =   "Form1.frx":1E5629
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   24
      Left            =   3720
      Picture         =   "Form1.frx":1F1AF5
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   23
      Left            =   10320
      Picture         =   "Form1.frx":1FDFC1
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   22
      Left            =   9000
      Picture         =   "Form1.frx":20A48D
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   21
      Left            =   7680
      Picture         =   "Form1.frx":216959
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   20
      Left            =   6360
      Picture         =   "Form1.frx":222E25
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   19
      Left            =   5040
      Picture         =   "Form1.frx":22F2F1
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image imgCover 
      Height          =   1575
      Index           =   18
      Left            =   3720
      Picture         =   "Form1.frx":23B7BD
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   29
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   28
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   27
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   26
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   25
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   24
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   23
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   22
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   21
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   20
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   19
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   18
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   6840
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   0
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   5
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   3
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   2
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   1
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   4
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   6
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   8
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   7
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   11
      Left            =   10320
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   10
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   9
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   15
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   14
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   13
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   12
      Left            =   3720
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1575
      Index           =   16
      Left            =   9000
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   4  'Dash-Dot
      Height          =   1215
      Left            =   8280
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3255
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   3615
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   2415
      Left            =   11640
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   2655
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H80000007&
      BackStyle       =   1  'Opaque
      Height          =   1215
      Left            =   3840
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   11880
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuMenu 
         Caption         =   "Menu"
      End
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clickMe As Variant
Dim a As Integer
Dim points As Integer

Private Sub Clears()
Me.lblPoints.Caption = "0"
Me.lblTries.Caption = "20"
End Sub

Private Sub Form_Load()
Call Clears
RandomizeMe
clickMe = 0
a = 1
Me.Text3.Text = clickMe
End Sub

Private Sub Command1_Click()
If MsgBox("Do you want to restart?", vbQuestion + vbYesNo, "Memory Game") = vbYes Then
    Call Form_Load
Else
Unload Me
frmMain.Show
End If
    For i = 0 To 29
    Me.imgCover(i).Visible = True
    Me.imgCover(i).Picture = Me.ImageList1.ListImages(1).Picture
Next
RandomizeMe
clickMe = 0
a = 1
Me.Text3.Text = clickMe
End Sub

Private Sub RandomizeMe()
Call Clears
Dim listIndex As Integer
Dim imageIndex As Integer
Dim size As Integer
LoadView
Me.ListView2.ListItems.Clear
size = 30
For i = 0 To 29
    listIndex = (Int(Rnd() * size)) + 1
    imageIndex = Val(Me.ListView1.ListItems(listIndex).Text)
    Me.Image1(i).Picture = Me.ImageList2.ListImages(imageIndex).Picture
    Set lst = Me.ListView2.ListItems.Add(, , i)
    lst.SubItems(1) = imageIndex
    Me.ListView1.ListItems.Remove (Me.ListView1.ListItems(listIndex).Index)
    size = size - 1
Next
End Sub

Private Sub LoadView()
For i = 1 To 30
    Set lst = Me.ListView1.ListItems.Add(, , i)
Next
End Sub

Private Sub imgCover_Click(Index As Integer)

clickMe = clickMe + 1
Me.Text3.Text = clickMe
If clickMe = 1 Then
    Me.imgCover(Index).Visible = False
    Me.Text1.Text = Index
ElseIf clickMe = 2 Then
    Me.imgCover(Index).Visible = False
    Me.Text2.Text = Index

ElseIf clickMe > 2 Then
    If Val(Me.ListView2.ListItems(Val(Me.Text1.Text) + 1).ListSubItems(1).Text) > Val(Me.ListView2.ListItems(Val(Me.Text2.Text) + 1).ListSubItems(1).Text) Then
        If Val(Val(Me.ListView2.ListItems(Val(Me.Text1.Text) + 1).ListSubItems(1).Text) - 15) = _
            Val(Me.ListView2.ListItems(Val(Me.Text2.Text) + 1).ListSubItems(1).Text) Then
            clickMe = 0
            Me.Text3.Text = clickMe
            Me.Text1.Text = ""
            Me.Text2.Text = ""
                 Call tries
                 Call score
        Else
            Me.imgCover(Val(Me.Text1.Text)).Visible = True
            Me.imgCover(Val(Me.Text2.Text)).Visible = True
            clickMe = 0
            Me.Text3.Text = clickMe
            Me.Text1.Text = ""
            Me.Text2.Text = ""
            Call tries
        End If

    ElseIf Val(Me.ListView2.ListItems(Val(Me.Text1.Text) + 1).ListSubItems(1).Text) < Val(Me.ListView2.ListItems(Val(Me.Text2.Text) + 1).ListSubItems(1).Text) Then
        If Val(Me.ListView2.ListItems(Val(Me.Text1.Text) + 1).ListSubItems(1).Text) = _
            Val(Val(Me.ListView2.ListItems(Val(Me.Text2.Text) + 1).ListSubItems(1).Text) - 15) Then
            clickMe = 0
                       
            Me.Text3.Text = clickMe
            Me.Text1.Text = ""
            Me.Text2.Text = ""
                    Call tries
                    Call score
        Else
            Me.imgCover(Val(Me.Text1.Text)).Visible = True
            Me.imgCover(Val(Me.Text2.Text)).Visible = True
            clickMe = 0
            Me.Text3.Text = clickMe
            Me.Text1.Text = ""
            Me.Text2.Text = ""
            Call tries
        End If

End If
End If

End Sub

Private Sub tries()
lblTries.Caption = lblTries.Caption - 1
End Sub

Private Sub score()
Me.lblPoints.Caption = Val(Me.lblPoints.Caption) + 2
End Sub

Private Sub lblPoints_Change()
If RS.State = adStateOpen Then RS.Close
RS.Open "SELECT*FROM Players WHERE Name LIKE '" & Me.Label5.Caption & "'", CN, adOpenStatic, adLockOptimistic
    With RS
        .Fields("Score") = Me.lblPoints.Caption
        .Update
        End With

If Me.lblPoints.Caption = "30" Then
MsgBox ("Game successful! Congrats :D"), vbInformation, "Memory Game"
Unload Me
frmMain.Show
End If
        
End Sub

Private Sub lblTries_Change()
If Me.lblTries.Caption = "15" Then
 MsgBox ("You only have 15 tries remaining! Be careful."), vbExclamation, "Memory Game"
End If
If Me.lblTries.Caption = "0" Then
MsgBox ("Your out of tries! "), vbExclamation, "Memory Game"
Call Command1_Click
End If
End Sub

Private Sub mnuMenu_Click()
Unload Me
frmMain.Show
End Sub

Private Sub Timer1_Timer()
On Error GoTo Err:
Me.imgTimer.Picture = LoadPicture(App.Path & "\Time\" & a & ".gif")
a = a + 1
If a = 122 Then
If MsgBox("Time is up, want to Play Again?", vbQuestion + vbYesNo, "Memory Game") = vbYes Then
    Form_Load
Else
Unload Me
frmMain.Show
End If
Exit Sub
Err:
a = 1
End If
End Sub
