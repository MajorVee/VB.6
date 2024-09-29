VERSION 5.00
Begin VB.Form frmExam 
   BackColor       =   &H00000000&
   Caption         =   "Final Exam"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13650
   LinkTopic       =   "Form2"
   ScaleHeight     =   8745
   ScaleWidth      =   13650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check20 
      BackColor       =   &H00000000&
      Caption         =   "Label"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   375
      Left            =   6840
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CheckBox Check19 
      BackColor       =   &H00000000&
      Caption         =   "Frame"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   375
      Left            =   6840
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CheckBox Check18 
      BackColor       =   &H00000000&
      Caption         =   "<=>"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   5040
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CheckBox Check17 
      BackColor       =   &H00000000&
      Caption         =   ">="
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   3480
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CheckBox Check16 
      BackColor       =   &H00000000&
      Caption         =   "<="
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1920
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CheckBox Check15 
      BackColor       =   &H00000000&
      Caption         =   "False"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1920
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CheckBox Check14 
      BackColor       =   &H00000000&
      Caption         =   "True"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CheckBox Check13 
      BackColor       =   &H00000000&
      Caption         =   "Charles Bobagge"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1800
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   6120
      Width           =   3375
   End
   Begin VB.CheckBox Check12 
      BackColor       =   &H00000000&
      Caption         =   "Charles Bobage"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   5520
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6120
      Width           =   3135
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H00000000&
      Caption         =   "Charles Bobbage"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF80FF&
      Height          =   375
      Left            =   8880
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   6120
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Re-Take"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      TabIndex        =   32
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CheckBox Check10 
      BackColor       =   &H00000000&
      Caption         =   "Frame"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CheckBox Check9 
      BackColor       =   &H00000000&
      Caption         =   "Textbox"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5160
      Width           =   2415
   End
   Begin VB.CheckBox Check8 
      BackColor       =   &H00000000&
      Caption         =   "Compile Error"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4080
      Width           =   2895
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00000000&
      Caption         =   "Logic Error"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CheckBox Check6 
      BackColor       =   &H00000000&
      Caption         =   "Methods"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CheckBox Check5 
      BackColor       =   &H00000000&
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00000000&
      Caption         =   "Methods"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00000000&
      Caption         =   "Properties"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Form"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00FFFF80&
      Caption         =   "Done :D"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "God Bless! :)"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8280
      TabIndex        =   51
      Top             =   240
      Width           =   1395
   End
   Begin VB.Label lblAns8 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Answer: >="
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   9000
      TabIndex        =   48
      Top             =   8280
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblWrong8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1080
      TabIndex        =   47
      Top             =   8160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8. Symbol for ""Greater than or Equal to"""
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   720
      TabIndex        =   44
      Top             =   7920
      Width           =   5400
   End
   Begin VB.Label lblAns7 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Answer: True"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   8760
      TabIndex        =   43
      Top             =   7440
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label lblWrong7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1080
      TabIndex        =   40
      Top             =   7320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7. Visual Basic is a High-Programming Language"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   720
      TabIndex        =   39
      Top             =   6960
      Width           =   6210
   End
   Begin VB.Label lblAns6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Answer: Charles Bobbage"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   6960
      TabIndex        =   38
      Top             =   6480
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.Label lblWrong6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1080
      TabIndex        =   37
      Top             =   6120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6. Also known as ""The Father of Computer"""
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   720
      TabIndex        =   33
      Top             =   5640
      Width           =   5535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Average:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   11625
      TabIndex        =   31
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblAve 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   615
      Left            =   11400
      TabIndex        =   30
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblItems 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "80"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1125
      Left            =   11775
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1125
      Left            =   11985
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label lblAns4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Answer: Logic Error"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   7920
      TabIndex        =   26
      Top             =   4080
      Visible         =   0   'False
      Width           =   3420
   End
   Begin VB.Label lblWrong5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1080
      TabIndex        =   25
      Top             =   5040
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblWrong4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1080
      TabIndex        =   24
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblWrong3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1080
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblWrong2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1080
      TabIndex        =   22
      Top             =   1920
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblAns5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Answer: Textbox"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   8520
      TabIndex        =   21
      Top             =   5160
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label lblAns3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Answer: Methods"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   8160
      TabIndex        =   20
      Top             =   3120
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Label lblAns2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Answer: Properties"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   7680
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   3240
   End
   Begin VB.Label lblWrong1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1080
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblAns1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Answer: Form"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   360
      Left            =   8640
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5. It allows the user to type an Input"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Top             =   4680
      Width           =   5130
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4. It is a type of Error; the project runs but produces incorrect results."
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   3600
      Width           =   9990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3. Action associated with the objects. Example: Print, Clear and Click"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Top             =   2640
      Width           =   9450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2. It tells something about the objects"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Top             =   1560
      Width           =   5265
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1. It is the windows and dialog boxes placed on the screen"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   7830
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "A. Check the box to choose the appropriate answer"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7905
   End
   Begin VB.Label lblBar 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "___"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1125
      Left            =   11520
      TabIndex        =   29
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackColor       =   &H00808080&
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   11040
      TabIndex        =   50
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmed by:Reyes :D
'March 5, 2016
'VB Exam

Private Sub Check1_Click()
Check2.Value = 0
Check19.Value = 0
End Sub

Private Sub Check10_Click()
Check9.Value = 0
Check20.Value = 0
End Sub

Private Sub Check11_Click()
Check12.Value = 0
Check13.Value = 0
End Sub

Private Sub Check12_Click()
Check13.Value = 0
Check11.Value = 0
End Sub

Private Sub Check13_Click()
Check12.Value = 0
Check11.Value = 0
End Sub

Private Sub Check16_Click()
Check17.Value = 0
Check18.Value = 0
End Sub

Private Sub Check17_Click()
Check16.Value = 0
Check18.Value = 0
End Sub

Private Sub Check18_Click()
Check17.Value = 0
Check16.Value = 0
End Sub

Private Sub Check19_Click()
Check1.Value = 0
Check2.Value = 0
End Sub

Private Sub Check2_Click()
Check1.Value = 0
Check19.Value = 0
End Sub

Private Sub Check20_Click()
Check9.Value = 0
Check10.Value = 0
End Sub

Private Sub Check3_Click()
Check4.Value = 0
End Sub

Private Sub Check4_Click()
Check3.Value = 0
End Sub

Private Sub Check5_Click()
Check6.Value = 0
End Sub

Private Sub Check6_Click()
Check5.Value = 0
End Sub

Private Sub Check7_Click()
Check8.Value = 0
End Sub

Private Sub Check8_Click()
Check7.Value = 0
End Sub

Private Sub Check9_Click()
Check10.Value = 0
Check20.Value = 0
End Sub

Private Sub cmdDone_Click()
    
If MsgBox("Are you sure you want to submit your answers?", vbQuestion + vbYesNo, "Final Exam Result") = vbNo Then
    frmExam.Show
    Exit Sub
    End If
 
'To show the Exam Results

Me.lblScore.Visible = True
Me.lblBar.Visible = True
Me.lblItems.Visible = True
Me.Label7.Visible = True
Me.Label11.Visible = True

'Shows the correct answers

If Me.Check2.Value = 0 Then
    Me.lblAns1.Visible = True
   Me.lblWrong1.Visible = True
    End If
    
If Me.Check3.Value = 0 Then
   Me.lblAns2.Visible = True
   Me.lblWrong2.Visible = True
    End If
    
If Me.Check6.Value = 0 Then
   Me.lblAns3.Visible = True
   Me.lblWrong3.Visible = True
    End If
    
If Me.Check7.Value = 0 Then
   Me.lblAns4.Visible = True
   Me.lblWrong4.Visible = True
    End If
    
If Me.Check9.Value = 0 Then
   Me.lblAns5.Visible = True
   Me.lblWrong5.Visible = True
    End If
    
If Check11.Value = 0 Then
    lblAns6.Visible = True
    lblWrong6.Visible = True
    End If
    
If Check14.Value = 0 Then
    Me.lblAns7.Visible = True
    lblWrong7.Visible = True
    End If
    
If Check17.Value = 0 Then
    lblAns8.Visible = True
    lblWrong8.Visible = True
    End If
    
    'SCORE
    
If Me.Check2.Value = 1 Then
    Me.lblScore.Caption = Val(Me.lblScore) + 10
    End If
    
If Me.Check3.Value = 1 Then
     Me.lblScore.Caption = Val(Me.lblScore) + 10
    End If
    
If Me.Check6.Value = 1 Then
      Me.lblScore.Caption = Val(Me.lblScore) + 10
    End If
    
If Me.Check7.Value = 1 Then
     Me.lblScore.Caption = Val(Me.lblScore) + 10
    End If
    
If Me.Check9.Value = 1 Then
    Me.lblScore.Caption = Val(Me.lblScore) + 10
    End If
    
If Check11.Value = 1 Then
    Me.lblScore.Caption = Val(Me.lblScore) + 10
    End If
    
If Check14.Value = 1 Then
    Me.lblScore.Caption = Val(Me.lblScore) + 10
    End If
    
If Check17.Value = 1 Then
    Me.lblScore.Caption = Val(Me.lblScore) + 10
    End If

    'Disables the Check Boxes
    
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
Check5.Enabled = False
Check5.Enabled = False
Check6.Enabled = False
Check7.Enabled = False
Check8.Enabled = False
Check9.Enabled = False
Check10.Enabled = False
Check11.Enabled = False
Check12.Enabled = False
Check13.Enabled = False
Check14.Enabled = False
Check15.Enabled = False
Check16.Enabled = False
Check17.Enabled = False
Check18.Enabled = False
Check19.Enabled = False
Check20.Enabled = False

    'Calculates the Average
    
lblAve.Caption = Val(Me.lblScore.Caption) / Val(Me.lblItems.Caption)
lblAve.Caption = Val(Me.lblAve.Caption) * 50
lblAve.Caption = Val(Me.lblAve.Caption) + 50 & " % "

    'Disables the Done button
    Me.cmdDone.Enabled = False

End Sub

Private Sub Command1_Click()
'Re-Take
Unload Me
frmExam.Show

'Disables the Re-Take Button
Me.Command1.Enabled = False
End Sub

Private Sub Command2_Click()
'Terminate the Program
End
End Sub

Private Sub lblScore_Change()
'Perfect? or Not?
If Val(Me.lblScore.Caption) = 100 Then
MsgBox "Perfect! :D Excellent!", vbInformation, "Final Exam Result"
Exit Sub
End If
End Sub
