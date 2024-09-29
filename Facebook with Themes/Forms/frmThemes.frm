VERSION 5.00
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#15.3#0"; "CODEJO~1.OCX"
Begin VB.Form frmThemes 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Facebook Profile"
   ClientHeight    =   10800
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   19530
   BeginProperty Font 
      Name            =   "Calibri Light"
      Size            =   12
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmThemes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13581.71
   ScaleMode       =   0  'User
   ScaleWidth      =   31173.92
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSlide 
      Interval        =   50
      Left            =   120
      Top             =   1080
   End
   Begin VB.Timer tmrSlideOff 
      Interval        =   10
      Left            =   720
      Top             =   1080
   End
   Begin VB.Frame MenuF 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   10815
      Left            =   0
      TabIndex        =   70
      Top             =   0
      Width           =   2925
      Begin VB.Label imgMenu 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M E N U"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   675
         Left            =   240
         TabIndex        =   73
         Top             =   2040
         Width           =   2175
      End
   End
   Begin VB.Frame MenuMenuF 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   10335
      Left            =   -240
      TabIndex        =   69
      Top             =   120
      Width           =   3165
      Begin VB.Image imgInfo 
         Height          =   10035
         Left            =   240
         Picture         =   "frmThemes.frx":000C
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   3135
      End
   End
   Begin VB.Frame Frame19 
      ForeColor       =   &H00FFC0C0&
      Height          =   2415
      Left            =   3000
      TabIndex        =   53
      Top             =   10080
      Width           =   16455
      Begin VB.Frame Frame22 
         Height          =   975
         Left            =   2400
         TabIndex        =   58
         Top             =   1320
         Width           =   2895
         Begin VB.CommandButton Command18 
            Caption         =   "Love"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   59
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image7 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":8E34
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame21 
         Height          =   975
         Left            =   6480
         TabIndex        =   56
         Top             =   1320
         Width           =   2895
         Begin VB.CommandButton Command17 
            Caption         =   "Comment"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   57
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image5 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":971A
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame20 
         Height          =   975
         Left            =   10560
         TabIndex        =   54
         Top             =   1320
         Width           =   2895
         Begin VB.CommandButton Command1 
            Caption         =   "Share"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   55
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image3 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":A7B9
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gahigugmaay mi sa akong Pamilya! HAHA"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         TabIndex        =   68
         Top             =   840
         Width           =   5940
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gilbert Apas and Bert L. Amija"
         BeginProperty Font 
            Name            =   "Eras Medium ITC"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   8400
         TabIndex        =   67
         Top             =   240
         Width           =   4755
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5 comments"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   13800
         TabIndex        =   66
         Top             =   1920
         Width           =   1410
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "28"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   65
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image Image35 
         Height          =   585
         Left            =   1200
         Picture         =   "frmThemes.frx":BD71
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   585
      End
      Begin VB.Image Image34 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   120
         Picture         =   "frmThemes.frx":C657
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bert  Apas"
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1680
         TabIndex        =   64
         Top             =   240
         Width           =   1725
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yesterday"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1800
         TabIndex        =   63
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         TabIndex        =   62
         Top             =   840
         Width           =   75
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "celebrating Valentines Day with"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   390
         Left            =   3600
         TabIndex        =   61
         Top             =   240
         Width           =   4035
      End
      Begin VB.Image Image33 
         Height          =   375
         Left            =   7680
         Picture         =   "frmThemes.frx":6E663
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Image8 
         Height          =   585
         Left            =   120
         Picture         =   "frmThemes.frx":6F559
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "42"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   720
         TabIndex        =   60
         Top             =   1920
         Width           =   360
      End
   End
   Begin VB.Frame Frame15 
      Height          =   2415
      Left            =   3000
      TabIndex        =   39
      Top             =   7560
      Width           =   16455
      Begin VB.Frame Frame18 
         Height          =   975
         Left            =   10560
         TabIndex        =   44
         Top             =   1320
         Width           =   2895
         Begin VB.CommandButton Command16 
            Caption         =   "Share"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   45
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image28 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":709D7
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame17 
         Height          =   975
         Left            =   6480
         TabIndex        =   42
         Top             =   1320
         Width           =   2895
         Begin VB.CommandButton Command15 
            Caption         =   "Comment"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   43
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image27 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":71F8F
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame16 
         Height          =   975
         Left            =   2400
         TabIndex        =   40
         Top             =   1320
         Width           =   2895
         Begin VB.CommandButton Command14 
            Caption         =   "L i k e"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   41
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image26 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":7302E
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "72"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   720
         TabIndex        =   52
         Top             =   1920
         Width           =   360
      End
      Begin VB.Image Image32 
         Height          =   585
         Left            =   1200
         Picture         =   "frmThemes.frx":744AC
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   585
      End
      Begin VB.Image Image31 
         Height          =   375
         Left            =   8280
         Picture         =   "frmThemes.frx":75531
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "extreme training with my Piggy"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   390
         Left            =   4200
         TabIndex        =   51
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I trained my Pig to Karate, now he's doing Pork Chops "
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         TabIndex        =   50
         Top             =   840
         Width           =   7680
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "14 Hours ago"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1800
         TabIndex        =   49
         Top             =   600
         Width           =   1290
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nam L. Reyes"
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1680
         TabIndex        =   48
         Top             =   240
         Width           =   2295
      End
      Begin VB.Image Image30 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   120
         Picture         =   "frmThemes.frx":76D3C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
      Begin VB.Image Image29 
         Height          =   585
         Left            =   120
         Picture         =   "frmThemes.frx":87338
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "14"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   47
         Top             =   1920
         Width           =   360
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7 comments"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   13800
         TabIndex        =   46
         Top             =   1920
         Width           =   1410
      End
   End
   Begin VB.Frame Frame11 
      Height          =   2175
      Left            =   3000
      TabIndex        =   23
      Top             =   5280
      Width           =   16455
      Begin VB.Frame Frame14 
         Height          =   975
         Left            =   2400
         TabIndex        =   28
         Top             =   1080
         Width           =   2895
         Begin VB.CommandButton Command13 
            Caption         =   "L i k e"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   29
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image20 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":88343
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame13 
         Height          =   975
         Left            =   6480
         TabIndex        =   26
         Top             =   1080
         Width           =   2895
         Begin VB.CommandButton Command12 
            Caption         =   "Comment"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   27
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image19 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":897C1
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame12 
         Height          =   975
         Left            =   10560
         TabIndex        =   24
         Top             =   1080
         Width           =   2895
         Begin VB.CommandButton Command11 
            Caption         =   "Share"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   25
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image18 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":8A860
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "13 comments"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   13920
         TabIndex        =   38
         Top             =   1680
         Width           =   1560
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "38"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   35
         Top             =   1680
         Width           =   360
      End
      Begin VB.Image Image22 
         Height          =   585
         Left            =   1200
         Picture         =   "frmThemes.frx":8BE18
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   585
      End
      Begin VB.Image Image24 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   120
         Picture         =   "frmThemes.frx":8CF96
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alvin Dulaugon"
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1680
         TabIndex        =   34
         Top             =   240
         Width           =   2505
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2 Hours ago"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1800
         TabIndex        =   33
         Top             =   600
         Width           =   1170
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "=============Loaded School Works!=============="
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         TabIndex        =   32
         Top             =   720
         Width           =   8070
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "feeling stressed"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   390
         Left            =   4320
         TabIndex        =   31
         Top             =   240
         Width           =   1965
      End
      Begin VB.Image Image23 
         Height          =   375
         Left            =   6360
         Picture         =   "frmThemes.frx":A2762
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Image Image21 
         Height          =   585
         Left            =   120
         Picture         =   "frmThemes.frx":A39C9
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "83"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   720
         TabIndex        =   30
         Top             =   1680
         Width           =   360
      End
   End
   Begin VB.Frame Frame7 
      Height          =   2295
      Left            =   3000
      TabIndex        =   11
      Top             =   2880
      Width           =   16455
      Begin VB.Frame Frame10 
         Height          =   975
         Left            =   10560
         TabIndex        =   20
         Top             =   1200
         Width           =   2895
         Begin VB.CommandButton Command10 
            Caption         =   "Share"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   21
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image16 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":A4E47
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame9 
         Height          =   975
         Left            =   6480
         TabIndex        =   18
         Top             =   1200
         Width           =   2895
         Begin VB.CommandButton Command9 
            Caption         =   "Comment"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   19
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image15 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":A63FF
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame8 
         Height          =   975
         Left            =   2400
         TabIndex        =   16
         Top             =   1200
         Width           =   2895
         Begin VB.CommandButton Command8 
            Caption         =   "L i k e"
            BeginProperty Font 
               Name            =   "Corbel"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   960
            TabIndex        =   17
            Top             =   270
            Width           =   1815
         End
         Begin VB.Image Image14 
            Height          =   705
            Left            =   120
            Picture         =   "frmThemes.frx":A749E
            Stretch         =   -1  'True
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3 comments"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   13800
         TabIndex        =   37
         Top             =   1800
         Width           =   1410
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   36
         Top             =   1680
         Width           =   180
      End
      Begin VB.Image Image25 
         Height          =   585
         Left            =   1200
         Picture         =   "frmThemes.frx":A891C
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   585
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   720
         TabIndex        =   22
         Top             =   1680
         Width           =   360
      End
      Begin VB.Image Image17 
         Height          =   585
         Left            =   120
         Picture         =   "frmThemes.frx":A9792
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   585
      End
      Begin VB.Image Image13 
         Height          =   495
         Left            =   10080
         Picture         =   "frmThemes.frx":AAC10
         Stretch         =   -1  'True
         Top             =   720
         Width           =   735
      End
      Begin VB.Image Image12 
         Height          =   375
         Left            =   7800
         Picture         =   "frmThemes.frx":ACB23
         Stretch         =   -1  'True
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "feeling determined"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   15.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   390
         Left            =   5280
         TabIndex        =   15
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Programming Language na mamaya. God Bless sa defense!"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   18
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1680
         TabIndex        =   14
         Top             =   720
         Width           =   8355
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8 minutes"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MC CS Programmers"
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   3420
      End
      Begin VB.Image Image11 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   120
         Picture         =   "frmThemes.frx":AE408
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   16455
      Begin VB.CommandButton Command7 
         Caption         =   "P  O  S  T"
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   14040
         TabIndex        =   9
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Frame Frame6 
         Height          =   975
         Left            =   10560
         TabIndex        =   7
         Top             =   1680
         Width           =   3375
         Begin VB.CommandButton Command6 
            Caption         =   "People"
            BeginProperty Font 
               Name            =   "Franklin Gothic Book"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Width           =   2055
         End
         Begin VB.Image Image10 
            Height          =   855
            Left            =   120
            Picture         =   "frmThemes.frx":F3D9B
            Stretch         =   -1  'True
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.Frame Frame4 
         Height          =   975
         Left            =   7080
         TabIndex        =   5
         Top             =   1680
         Width           =   3375
         Begin VB.CommandButton Command4 
            Caption         =   "Check In"
            BeginProperty Font 
               Name            =   "Franklin Gothic Book"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1200
            TabIndex        =   6
            Top             =   240
            Width           =   2055
         End
         Begin VB.Image Image6 
            Height          =   855
            Left            =   120
            Picture         =   "frmThemes.frx":F4ECE
            Stretch         =   -1  'True
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Height          =   975
         Left            =   3600
         TabIndex        =   3
         Top             =   1680
         Width           =   3375
         Begin VB.CommandButton Command3 
            Caption         =   "Photo"
            BeginProperty Font 
               Name            =   "Franklin Gothic Book"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1200
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
         Begin VB.Image Image4 
            Appearance      =   0  'Flat
            Height          =   855
            Left            =   120
            Picture         =   "frmThemes.frx":F7514
            Stretch         =   -1  'True
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   3255
         Begin VB.CommandButton Command2 
            Caption         =   "Live"
            BeginProperty Font 
               Name            =   "Franklin Gothic Book"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   1080
            MaskColor       =   &H00C0C0FF&
            Picture         =   "frmThemes.frx":105DC0
            TabIndex        =   2
            Top             =   240
            Width           =   2055
         End
         Begin VB.Image Image1 
            Height          =   855
            Left            =   120
            Picture         =   "frmThemes.frx":106536
            Stretch         =   -1  'True
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What's on your mind?"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   24
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   4185
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   120
         Picture         =   "frmThemes.frx":10A908
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label lblSlideCount 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   72
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label lblSlideSwicth 
      Caption         =   "Label1"
      Height          =   255
      Left            =   480
      TabIndex        =   71
      Top             =   600
      Width           =   2535
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework1 
      Left            =   4680
      Top             =   5160
      _Version        =   983043
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu mnuAS 
      Caption         =   "Account Settings"
      Begin VB.Menu mnuClose 
         Caption         =   "Logout"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuCF 
         Caption         =   "Close Facebook"
      End
   End
End
Attribute VB_Name = "frmThemes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuCF_Click()
End
End Sub

Private Sub mnuClose_Click()
Unload Me
frmLogin.Show
End Sub
Private Sub Form_Load()
Me.lblSlideCount.Caption = "1"
Me.lblSlideSwicth.Caption = "OFF"
Me.tmrSlide.Enabled = False
Me.tmrSlideOff.Enabled = False
Me.lblSlideCount.Visible = False
Me.lblSlideSwicth.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.lblSlideSwicth.Caption = "OFF" Then
Exit Sub
Else
Me.MenuMenuF.Left = 0
    Me.tmrSlideOff.Enabled = True
    Me.lblSlideSwicth.Caption = "OFF"
End If
End Sub

Private Sub imgMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.lblSlideSwicth.Caption = "ONMENU" Then
Exit Sub
Else
    Me.tmrSlide.Enabled = True
    Me.lblSlideSwicth.Caption = "ONMENU"
End If
End Sub

Private Sub MenuF_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Me.lblSlideSwicth.Caption = "OFF" Then
Exit Sub
Else
Me.MenuMenuF.Left = 0
    Me.tmrSlideOff.Enabled = True
    Me.lblSlideSwicth.Caption = "OFF"
End If
End Sub

Private Sub tmrSlide_Timer()
If Me.lblSlideSwicth.Caption = "ONMENU" Then
    If Me.lblSlideCount.Caption = "1" Then
    Me.MenuMenuF.Left = MenuMenuF.Left + 2500
    Me.lblSlideCount.Caption = Val(Me.lblSlideCount.Caption) + 1
    ElseIf Me.lblSlideCount.Caption = "2" Then
    Me.MenuMenuF.Left = MenuMenuF.Left + 2500
    Me.lblSlideCount.Caption = Val(Me.lblSlideCount.Caption) + 1
    ElseIf Me.lblSlideCount.Caption = "3" Then
    Me.MenuMenuF.Left = MenuMenuF.Left + 135
    Me.lblSlideCount.Caption = Val(Me.lblSlideCount.Caption) + 1
    ElseIf Me.lblSlideCount.Caption = "4" Then
    GoTo StopBaby:
    Else
    Exit Sub
    End If
Else
Exit Sub
End If
Exit Sub
StopBaby:
        Me.tmrSlide.Enabled = False
        Me.lblSlideCount.Caption = "1"
End Sub

Private Sub tmrSlideOff_Timer()
If Me.lblSlideSwicth.Caption = "ONMENU" Then
    Me.MenuMenuF.Left = 0
    GoTo StopBaby:
    Exit Sub
Else
Exit Sub
End If
Exit Sub
StopBaby:
        Me.lblSlideSwicth.Caption = "OFF"
        Me.tmrSlideOff.Enabled = False
        Me.lblSlideCount.Caption = "1"
End Sub

