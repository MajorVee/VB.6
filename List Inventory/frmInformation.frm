VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInformation 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fruit Sales"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10815
   ControlBox      =   0   'False
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
   ScaleHeight     =   8475
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   8070
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12632064
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Fruit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Weight"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Order"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Contact"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H000000FF&
      Caption         =   "X"
      Height          =   615
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H0080FF80&
      Caption         =   "PRINT PREVIEW"
      Height          =   615
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000080FF&
      Caption         =   "DELETE"
      Height          =   615
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H8000000D&
      Caption         =   "REFRESH"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H0000FFFF&
      Caption         =   "SAVE"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7440
      Width           =   2055
   End
   Begin VB.ComboBox cmbOrder 
      Height          =   435
      Left            =   6600
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox cmbSize 
      Height          =   435
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtContact 
      Height          =   435
      Left            =   6600
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtWeight 
      Height          =   435
      Left            =   6600
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtFruit 
      Height          =   435
      Left            =   2640
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtID 
      Height          =   435
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FRUIT SALES INVENTORY"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   540
      Left            =   3000
      TabIndex        =   17
      Top             =   120
      Width           =   4920
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORDER"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   5400
      TabIndex        =   5
      Top             =   1440
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WEIGHT"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      Top             =   840
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   5040
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIZE"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FRUIT"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ID"
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "frmInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
If MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Fruit Sales Inventory System") = vbYes Then
    End
End If
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Delete " & Me.ListView1.SelectedItem.Text & "?", vbQuestion + vbYesNo, "Fruit Sales Inventory System") = vbYes Then
CN.Execute "delete from Information WHERE ID LIKE'" & Me.txtID.Text & "'"
cmdRefresh_Click
Exit Sub
Else
End If
End Sub

Private Sub cmdPrint_Click()
If Me.ListView1.ListItems.Count = 0 Then Exit Sub
Set xcl = Excel.Application
xcl.Workbooks.Open (App.Path & "\Sample.xlsx"), , True
xcl.SheetsInNewWorkbook = 1
Row = 5

For i = 1 To Me.ListView1.ListItems.Count
xcl.ActiveSheet.Cells(Row, 2).Value = Me.ListView1.ListItems(i).ListSubItems(1).Text
xcl.ActiveSheet.Cells(Row, 3).Value = Me.ListView1.ListItems(i).ListSubItems(2).Text
xcl.ActiveSheet.Cells(Row, 4).Value = Me.ListView1.ListItems(i).ListSubItems(3).Text
xcl.ActiveSheet.Cells(Row, 5).Value = Me.ListView1.ListItems(i).ListSubItems(4).Text
xcl.ActiveSheet.Cells(Row, 6).Value = Me.ListView1.ListItems(i).ListSubItems(5).Text

Row = Row + 1
Next i
xcl.Visible = True
xcl.Application.ActiveSheet.PrintPreview
End Sub

Private Sub cmdRefresh_Click()
txtID.Text = "(NEW)"
txtContact.Text = ""
txtFruit.Text = ""
txtWeight.Text = ""
cmbOrder.Clear
cmbSize.Clear

Call LoadView
Call LoadSize
Call LoadOrder

cmdSave.Caption = "SAVE"
cmdDelete.Enabled = False

End Sub

Private Sub LoadSize()
If RS.State = adStateOpen Then RS.Close
RS.Open "SELECT*FROM SetSize", CN, adOpenStatic, adLockOptimistic
If RS.RecordCount = 0 Then Exit Sub
While Not RS.EOF
Me.cmbSize.AddItem RS("Size")
RS.MoveNext
Wend

End Sub

Private Sub LoadOrder()
If RS.State = adStateOpen Then RS.Close
RS.Open "SELECT*FROM SetOrder", CN, adOpenStatic, adLockOptimistic
If RS.RecordCount = 0 Then Exit Sub
While Not RS.EOF
Me.cmbOrder.AddItem RS("Order")
RS.MoveNext
Wend

End Sub

Private Sub cmdSave_Click()
If Me.txtFruit.Text = "" Then
MsgBox "Please type your desired Fruit", vbExclamation, "Fruit Order"
Me.txtFruit.SetFocus
Exit Sub
End If
If Me.txtWeight.Text = "" Then
MsgBox "Please type the desired Weight for your Order", vbExclamation, "Fruit Order"
Me.txtWeight.SetFocus
Exit Sub
End If
If Me.txtContact.Text = "" Then
MsgBox "Please fill the Contact", vbExclamation, "Fruit Order"
Me.txtContact.SetFocus
Exit Sub
End If


If Me.cmdSave.Caption = "SAVE" Then
If RS.State = adStateOpen Then RS.Close

RS.Open "SELECT*FROM Information", CN, adOpenStatic, adLockOptimistic
With RS
    .AddNew
    .Fields("Fruit") = txtFruit.Text
    .Fields("Size") = cmbSize.Text
    .Fields("Weight") = txtWeight.Text
    .Fields("Order") = cmbOrder.Text
    .Fields("Contact") = txtContact.Text
    
    .Update
    
End With
MsgBox "Successfully SAVED! :D", vbInformation
cmdRefresh_Click
Exit Sub
Else

If RS.State = adStateOpen Then RS.Close
RS.Open "SELECT*FROM Information WHERE ID LIKE'" & Me.txtID.Text & "'", CN, adOpenStatic, adLockOptimistic
With RS
    .Fields("Fruit") = txtFruit.Text
    .Fields("Size") = cmbSize.Text
    .Fields("Weight") = txtWeight.Text
    .Fields("Order") = cmbOrder.Text
    .Fields("Contact") = txtContact.Text
    
    .Update
End With
MsgBox "Successfully UPDATED! :D", vbInformation, "Fruit Order"
cmdRefresh_Click
End If
End Sub

Private Sub Form_Load()
cmdRefresh_Click
End Sub

Private Sub LoadView()
If RS.State = adStateOpen Then RS.Close
RS.Open "SELECT*FROM Information", CN, adOpenStatic, adLockOptimistic
Me.ListView1.ListItems.Clear
If RS.RecordCount = 0 Then Exit Sub
While Not RS.EOF
Set lst = Me.ListView1.ListItems.Add(, , RS("ID"))
lst.SubItems(1) = RS("Fruit")
lst.SubItems(2) = RS("Size")
lst.SubItems(3) = RS("Weight")
lst.SubItems(4) = RS("Order")
lst.SubItems(5) = RS("Contact")
RS.MoveNext
Wend
End Sub

Private Sub ListView1_Click()
If Me.ListView1.ListItems.Count = 0 Then Exit Sub
If RS.State = adStateOpen Then RS.Close

RS.Open "SELECT*FROM Information WHERE ID LIKE'" & Me.ListView1.SelectedItem.Text & "'", CN, adOpenStatic, adLockOptimistic

Me.txtID.Text = RS("ID")
Me.txtFruit.Text = RS("Fruit")
Me.cmbSize.Text = RS("Size")
Me.txtWeight.Text = RS("Weight")
Me.cmbOrder.Text = RS("Order")
Me.txtContact.Text = RS("Contact")

Me.cmdDelete.Enabled = True
Me.cmdSave.Caption = "UPDATE"
End Sub

Private Sub txtWeight_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    KeyAscii = 8
ElseIf Not Chr(KeyAscii) Like "[0-9-.]" Then
    KeyAscii = 0
    MsgBox "Please use in a digit form.", vbInformation, "Fruit Order"
    Me.txtWeight.Text = ""
End If
End Sub
