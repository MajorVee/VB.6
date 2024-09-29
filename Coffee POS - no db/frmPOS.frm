VERSION 5.00
Begin VB.Form frmPOS 
   BackColor       =   &H00004080&
   Caption         =   "Kape-Libro"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11265
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   11265
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame7 
      BackColor       =   &H00004080&
      Caption         =   "Food"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   3375
      Left            =   4800
      TabIndex        =   31
      Top             =   4920
      Width           =   2775
      Begin VB.OptionButton optTruffle 
         BackColor       =   &H00004080&
         Caption         =   "Choco Truffle"
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
         Height          =   255
         Left            =   360
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   34
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton optPop 
         BackColor       =   &H00004080&
         Caption         =   "Iced Coffee Popsicle"
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
         Height          =   795
         Left            =   360
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   33
         Top             =   2040
         Width           =   2055
      End
      Begin VB.OptionButton optTruffleM 
         BackColor       =   &H00004080&
         Caption         =   "Mocha Truffle"
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
         Height          =   540
         Left            =   360
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   32
         Top             =   600
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Back"
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
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H000000FF&
      Caption         =   "Close Shop"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00004080&
      Caption         =   "Coffee"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   4800
      TabIndex        =   22
      Top             =   1200
      Width           =   2775
      Begin VB.OptionButton optNut 
         BackColor       =   &H00004080&
         Caption         =   "Nutella Hazelnut"
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
         Height          =   555
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   28
         Top             =   2880
         Width           =   2415
      End
      Begin VB.OptionButton optFrappucino 
         BackColor       =   &H00004080&
         Caption         =   "Frappuccino"
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
         Height          =   315
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   27
         Top             =   2400
         Width           =   2295
      End
      Begin VB.OptionButton optIced 
         BackColor       =   &H00004080&
         Caption         =   "Iced Cappuccino"
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
         Height          =   690
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   26
         Top             =   840
         Width           =   2295
      End
      Begin VB.OptionButton optLatte 
         BackColor       =   &H00004080&
         Caption         =   "Latte"
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
         Height          =   255
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   25
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton optEspresso 
         BackColor       =   &H00004080&
         Caption         =   "Espresso"
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
         Height          =   255
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   24
         Top             =   1920
         Width           =   2295
      End
      Begin VB.OptionButton optCappucino 
         BackColor       =   &H00004080&
         Caption         =   "Cappuccino"
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
         Height          =   300
         Left            =   240
         MaskColor       =   &H00C0C0FF&
         TabIndex        =   23
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "STATUS:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   2775
      Left            =   7680
      TabIndex        =   17
      Top             =   1200
      Width           =   3495
      Begin VB.Label lblTotalSales 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   21
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Sales:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1080
         TabIndex        =   20
         Top             =   1560
         Width           =   1320
      End
      Begin VB.Label lblNoOrders 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   19
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Orders:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   840
         TabIndex        =   18
         Top             =   480
         Width           =   2130
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000C0&
      Height          =   7215
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4575
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         TabIndex        =   15
         Top             =   5520
         Width           =   2655
         Begin VB.Label lblTotal 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFC0&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1260
            TabIndex        =   16
            Top             =   480
            Width           =   315
         End
      End
      Begin VB.CommandButton cmdNew 
         BackColor       =   &H00C0FFC0&
         Caption         =   "New Order!"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   360
         TabIndex        =   9
         Top             =   3240
         Width           =   3735
         Begin VB.Label lblTax 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2760
            TabIndex        =   13
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Take-Out Tax:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   240
            TabIndex        =   12
            Top             =   1320
            Width           =   1965
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   2760
            TabIndex        =   11
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   720
            TabIndex        =   10
            Top             =   600
            Width           =   1395
         End
      End
      Begin VB.CommandButton cmdCalculate 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Calculate the Order :D"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmdClear 
         BackColor       =   &H0080C0FF&
         Caption         =   "Clear for new order :D"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.CheckBox chkTax 
         BackColor       =   &H00000080&
         Caption         =   "Take-Out"
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
         Height          =   405
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   390
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lblItemAmount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2400
         TabIndex        =   5
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   1560
         Width           =   990
      End
   End
   Begin VB.Image Image3 
      Height          =   4260
      Left            =   7920
      Picture         =   "frmPOS.frx":0000
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   3180
   End
   Begin VB.Image Image2 
      Height          =   840
      Left            =   10200
      Picture         =   "frmPOS.frx":1E28E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   1275
      Left            =   120
      Picture         =   "frmPOS.frx":257FC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2235
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kape-Libro Coffee Shop"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   840
      Left            =   2400
      TabIndex        =   0
      Top             =   120
      Width           =   7845
   End
End
Attribute VB_Name = "frmPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmed by: Nyork Sardua -.- & Namtot :}
'Date:Feb. 05, 2016
'A Simple POS without a Database
'Kape-Libro Coffee Shop

Dim mcSubTotal As Currency
Dim mcTotal As Currency
Dim mcGrandTotal As Currency
Dim miCustomerCount As Integer
Const cCAPPUCINO_PRICE As Currency = 8
Const cESPRESSO_PRICE As Currency = 6
Const cLATTE_PRICE As Currency = 7
Const cICED_PRICE As Currency = 9
Const cFRAPPUCINO_PRICE As Currency = 3
Const cTRUFFLEM_PRICE As Currency = 4.5
Const cTRUFFLE_Price As Currency = 4
Const cPOP_PRICE As Currency = 5
Const cNUT_Price As Currency = 10

Private Sub cmdBack_Click()
Unload Me
frmOpen.Show
End Sub

Private Sub cmdCalculate_Click()
'Calculates all the amount
Dim cPrice As Currency
Dim iQuantity As Integer
Dim cTax As Currency
Dim cItemAmount As Currency
Const cTAX_RATE As Currency = 0.08

'For the Price
If optCappucino.Value = True Then
    cPrice = cCAPPUCINO_PRICE
ElseIf optEspresso.Value = True Then
    cPrice = cESPRESSO_PRICE
ElseIf optLatte.Value = True Then
    cPrice = cLATTE_PRICE
ElseIf optIced.Value = True Then
    cPrice = cICED_PRICE
ElseIf optFrappucino.Value = True Then
    cPrice = cFRAPPUCINO_PRICE
ElseIf optPop.Value = True Then
    cPrice = cPOP_PRICE
ElseIf optTruffleM.Value = True Then
    cPrice = cTRUFFLEM_PRICE
ElseIf optTruffle.Value = True Then
    cPrice = cTRUFFLE_Price
ElseIf optNut.Value = True Then
    cPrice = cNUT_Price
Else
    MsgBox "Please choose an Order!", vbExclamation, "Kape-Libro Coffee Shop"
End If

'Add the price x quantity to price so far
If IsNumeric(txtQuantity.Text) Then
    iQuantity = Val(txtQuantity.Text)
    cItemAmount = cPrice * iQuantity
    mcSubTotal = mcSubTotal + cItemAmount
    
        If chkTax.Value = Checked Then
    cTax = mcSubTotal * cTAX_RATE
    
End If
    
    mcTotal = mcSubTotal + cTax
    lblItemAmount.Caption = "$" & Format(cItemAmount, "standard")
    lblSubTotal.Caption = "$" & Format(mcSubTotal, "standard")
    lblTax.Caption = "$" & Format(cTax, "standard")
    lblTotal.Caption = "$" & Format(mcTotal, "standard")
    
Else
        MsgBox "Quantity must be Numeric", vbExclamation, "Kape-Libro Coffee Shop"
        txtQuantity.SetFocus
        
        End If
        
End Sub

Private Sub cmdClear_Click()
'To empty the quantity
txtQuantity = ""

'For Tax
chkTax.Enabled = False
With txtQuantity
    .Text = ""
    .SetFocus
End With

'For Option Button
optCappucino.Value = True
optLatte.Value = False
optIced.Value = False
optEspresso.Value = False
optFrappucino.Value = False
optTruffle.Value = False
optTruffleM.Value = False
optPop.Value = False
optNut.Value = False

'For Amount, SubTotal, Tax
lblItemAmount.Caption = ""
lblSubTotal.Caption = ""
End Sub

Private Sub cmdClose_Click()
'Terminate the Program :(
If MsgBox("Are you sure you want to close the Shop?", vbQuestion + vbYesNo, "Kape-Libro Coffee Shop") = vbYes Then
    End
End If
End Sub

Private Sub cmdNew_Click()
'To empty the quantity
txtQuantity = ""

'To empty the Tax
chkTax = Unchecked

'To empty the Option Button
optLatte = False
optCappucino = False
optIced = False
optExpresso = False
optFrappucino = False
optTruffle.Value = False
optTruffleM.Value = False
optPop.Value = False
optNut.Value = False

'To empty the Amount, SubTotal, Tax, Total
lblItemAmount.Caption = 0
lblSubTotal.Caption = 0
lblTax.Caption = 0
lblTotal.Caption = 0

'Add to Totals
mcGrandTotal = mcGrandTotal + mcTotal
mcSubTotal = 0
mcTotal = 0 'to reset for new costumer
miCustomerCount = miCustomerCount + 1

'Enable's checkbox
With chkTax
.Enabled = True
.Value = False
End With

'Displays the total sales and and total count of orders
lblNoOrders.Caption = miCustomerCount
lblTotalSales.Caption = "$" & Format(mcGrandTotal, "standard")
End Sub
