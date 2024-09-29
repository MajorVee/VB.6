Attribute VB_Name = "modFunctions"

Option Explicit
Public client As String
Public car As String
Public driver As String
Public Trans As String

Public Sub UserMan_Clear()
With frmUser
    .txtUsername.Text = vbNullString
    .txtPassword.Text = vbNullString
    .txtConfirm.Text = vbNullString
    .txtName.Text = vbNullString
'    .cboRole.Text = vbNullString
    .dtDate.Value = Date
    .chkPass.Value = "0"
End With
End Sub

Public Sub RefAccount()
    frmUser.lvUserAccount.ListItems.Clear
    Set rs = New ADODB.Recordset
    rs.Open "Select * from [User] Order by UserID", conn, adOpenForwardOnly, adLockPessimistic
    With rs
            Do While Not .EOF
            'With frmUser
                frmUser.lvUserAccount.ListItems.Add , , !UserID
                                                                                                        'tinanggal ko yung ,1, 1 eto yung sa icon  sa ID
                frmUser.lvUserAccount.ListItems(frmUser.lvUserAccount.ListItems.Count).SubItems(1) = "" & !UserName
                frmUser.lvUserAccount.ListItems(frmUser.lvUserAccount.ListItems.Count).SubItems(2) = "" & !CompName
                frmUser.lvUserAccount.ListItems(frmUser.lvUserAccount.ListItems.Count).SubItems(3) = "" & !role
                frmUser.lvUserAccount.ListItems(frmUser.lvUserAccount.ListItems.Count).SubItems(4) = "" & !DateReg
                frmUser.lvUserAccount.ListItems(frmUser.lvUserAccount.ListItems.Count).SubItems(5) = "" & !Password
                frmUser.lvUserAccount.ListItems(frmUser.lvUserAccount.ListItems.Count).SubItems(6) = "" & !Confirm
                frmUser.lvUserAccount.ListItems(frmUser.lvUserAccount.ListItems.Count).SubItems(7) = "" & !Email
            'End With
            .MoveNext
            Loop
            .Close
    End With
Set rs = Nothing
                frmUser.lvUserAccount.ColumnHeaders(6).Width = 0
                frmUser.lvUserAccount.ColumnHeaders(7).Width = 0
                frmUser.lvUserAccount.ColumnHeaders(8).Width = 0
End Sub

Public Sub Client_Clear()
With frmClient
    .txtFirstname.Text = vbNullString
    .txtMidname.Text = vbNullString
    .txtLastname.Text = vbNullString
    .txtAddress.Text = vbNullString
    .dtBirth.Value = Date
    .cboGender.Text = vbNullString
    .txtOccup.Text = vbNullString
    .txtPhone.Text = vbNullString
    .txtEmail.Text = vbNullString
    .txtLicense.Text = vbNullString
    .dtDateReg.Value = Date
End With
End Sub

Public Sub Client_Lock()
With frmClient
    .txtFirstname.Enabled = False
    .txtMidname.Enabled = False
    .txtLastname.Enabled = False
    .txtAddress.Enabled = False
    .dtBirth.Enabled = False
    .cboGender.Enabled = False
    .txtOccup.Enabled = False
    .txtPhone.Enabled = False
    .txtEmail.Enabled = False
    .txtLicense.Enabled = False
    .dtDateReg.Enabled = False
End With
End Sub
'
Public Sub Client_Unlock()
With frmClient
    .txtFirstname.Enabled = True
    .txtMidname.Enabled = True
    .txtLastname.Enabled = True
    .txtAddress.Enabled = True
    .dtBirth.Enabled = True
    .cboGender.Enabled = True
    .txtOccup.Enabled = True
    .txtPhone.Enabled = True
    .txtEmail.Enabled = True
    .txtLicense.Enabled = True
    .dtDateReg.Enabled = True
    .txtFirstname.SetFocus
End With
End Sub

Public Sub RefClient()
    frmClient.lvClient.ListItems.Clear
    Set rs = New ADODB.Recordset
    rs.Open "Select * from Client Order by ClientID", conn, adOpenForwardOnly, adLockPessimistic
    With rs
   
            Do While Not .EOF
            'With frmUserManager
                frmClient.lvClient.ListItems.Add , , !ClientID
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(1) = "" & !FName
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(2) = "" & !MName
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(3) = "" & !LName
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(4) = "" & !Address
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(5) = "" & !DOB
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(6) = "" & !Gender
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(7) = "" & !Occup
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(8) = "" & !PhoneNo
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(9) = "" & !Email
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(10) = "" & !License
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(11) = "" & !DateReg
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(12) = "" & !Postal
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(13) = "" & !Company
                frmClient.lvClient.ListItems(frmClient.lvClient.ListItems.Count).SubItems(14) = "" & !Clearance
            'End With
            .MoveNext
            Loop
            .Close
    End With
Set rs = Nothing
            Dim i%
                frmClient.lvClient.ColumnHeaders(3).Width = 0
                frmClient.lvClient.ColumnHeaders(5).Width = 0
                frmClient.lvClient.ColumnHeaders(6).Width = 0
                frmClient.lvClient.ColumnHeaders(9).Width = 0
                frmClient.lvClient.ColumnHeaders(10).Width = 0
                frmClient.lvClient.ColumnHeaders(11).Width = 0
                frmClient.lvClient.ColumnHeaders(13).Width = 0
                frmClient.lvClient.ColumnHeaders(14).Width = 0
                frmClient.lvClient.ColumnHeaders(15).Width = 0
End Sub
Public Sub NoMaleClient()
Set rs = New ADODB.Recordset
rs.Open "Select * from Client where Gender='" & "Male" & "'", conn, adOpenForwardOnly, adLockPessimistic
frmClient.Label26.Caption = rs.RecordCount
rs.Close
Set rs = Nothing
End Sub
Public Sub NoFemaleClient()
Set rs = New ADODB.Recordset
rs.Open "Select * from Client where Gender='" & "Female" & "'", conn, adOpenForwardOnly, adLockPessimistic
frmClient.Label24.Caption = rs.RecordCount
rs.Close
Set rs = Nothing
End Sub

Public Sub Car_Clear()
With frmCar
    .txtPlate.Text = vbNullString
    .txtReg.Text = vbNullString
    .txtType.Text = vbNullString
    .dtManu.Value = Date
    .txtModel.Text = vbNullString
    .cboColor.Text = vbNullString
    .cboMake.Text = vbNullString
    .txtSpeed.Text = vbNullString
    .cboCondition.Text = vbNullString
    .cboStatus.Text = vbNullString
End With
End Sub

Public Sub Car_Lock()
With frmCar
    .txtPlate.Enabled = False
    .txtReg.Enabled = False
    .txtType.Enabled = False
    .dtManu.Enabled = False
    .txtModel.Enabled = False
    .cboColor.Enabled = False
    .cboMake.Enabled = False
    .txtSpeed.Enabled = False
    .cboCondition.Enabled = False
    .cboStatus.Enabled = False
End With
End Sub

Public Sub Car_Unlock()
With frmCar
    .txtPlate.Enabled = True
    .txtReg.Enabled = True
    .txtType.Enabled = True
    .dtManu.Enabled = True
    .txtModel.Enabled = True
    .cboColor.Enabled = True
    .cboMake.Enabled = True
    .txtSpeed.Enabled = True
    .cboCondition.Enabled = True
    .cboStatus.Enabled = True
    .txtPlate.SetFocus
End With
End Sub
Public Sub RefCar()
    frmCar.lvCars.ListItems.Clear
    Set rs = New ADODB.Recordset
    rs.Open "Select * from Car Order by CarID", conn, adOpenForwardOnly, adLockPessimistic
    With rs
            Do While Not .EOF
            'With frmUser
                                                       
                frmCar.lvCars.ListItems.Add , , !CarID, 1, 1
                frmCar.lvCars.ListItems(frmCar.lvCars.ListItems.Count).SubItems(1) = "" & !PlateNo
                frmCar.lvCars.ListItems(frmCar.lvCars.ListItems.Count).SubItems(2) = "" & !RegNo
                frmCar.lvCars.ListItems(frmCar.lvCars.ListItems.Count).SubItems(3) = "" & !Type
                frmCar.lvCars.ListItems(frmCar.lvCars.ListItems.Count).SubItems(4) = "" & !Trip
                frmCar.lvCars.ListItems(frmCar.lvCars.ListItems.Count).SubItems(5) = "" & !DateManu
                frmCar.lvCars.ListItems(frmCar.lvCars.ListItems.Count).SubItems(6) = "" & !Model
                frmCar.lvCars.ListItems(frmCar.lvCars.ListItems.Count).SubItems(7) = "" & !Make
                frmCar.lvCars.ListItems(frmCar.lvCars.ListItems.Count).SubItems(8) = "" & !Speed
                frmCar.lvCars.ListItems(frmCar.lvCars.ListItems.Count).SubItems(9) = "" & !Condition
                frmCar.lvCars.ListItems(frmCar.lvCars.ListItems.Count).SubItems(10) = "" & !AvailStatus

            'End With
            .MoveNext
            Loop
            .Close
    End With
Set rs = Nothing
                frmCar.lvCars.ColumnHeaders(3).Width = 0
                frmCar.lvCars.ColumnHeaders(4).Width = 0
                frmCar.lvCars.ColumnHeaders(6).Width = 0
                frmCar.lvCars.ColumnHeaders(7).Width = 0
                frmCar.lvCars.ColumnHeaders(9).Width = 0
                frmCar.lvCars.ColumnHeaders(10).Width = 0
End Sub

Public Sub NoCarAvail()
Set rs = New ADODB.Recordset
rs.Open "Select * from Car where AvailStatus='" & "Available" & "'", conn, adOpenForwardOnly, adLockPessimistic
frmCar.Label27.Caption = rs.RecordCount
rs.Close
Set rs = Nothing
End Sub

Public Sub NoCarCondition()
Set rs = New ADODB.Recordset
rs.Open "Select * from Car where Condition='" & "Excellent" & "'", conn, adOpenForwardOnly, adLockPessimistic
frmCar.Label33.Caption = rs.RecordCount
rs.Close
Set rs = Nothing
End Sub

Public Sub PopColor()
Set rs = New ADODB.Recordset
rs.Open "Select Trip from Trip Order by Trip ASC", conn, 3, 3
Do While Not rs.EOF
    frmCar.cboColor.AddItem rs!Trip
rs.MoveNext
Loop
End Sub

Public Sub PopMake()
Set rs = New ADODB.Recordset
rs.Open "Select Brand from Make Order by Brand ASC", conn, 3, 3
Do While Not rs.EOF
    frmCar.cboMake.AddItem rs!Brand
rs.MoveNext
Loop

End Sub

Public Sub RefLog()
    frmLog.lvLog.ListItems.Clear
    Set rs = New ADODB.Recordset
    rs.Open "Select * from Logs Order by LogID DESC", conn, adOpenForwardOnly, adLockPessimistic
    With rs
            Do While Not .EOF
                frmLog.lvLog.ListItems.Add , , !LogID, 1, 1
                frmLog.lvLog.ListItems(frmLog.lvLog.ListItems.Count).SubItems(1) = "" & !UserName
                frmLog.lvLog.ListItems(frmLog.lvLog.ListItems.Count).SubItems(2) = "" & !CompName
                frmLog.lvLog.ListItems(frmLog.lvLog.ListItems.Count).SubItems(3) = "" & !TimeIn
                frmLog.lvLog.ListItems(frmLog.lvLog.ListItems.Count).SubItems(4) = "" & !TimeOut
                frmLog.lvLog.ListItems(frmLog.lvLog.ListItems.Count).SubItems(5) = "" & !LogDate
            .MoveNext
            Loop
            .Close
    End With
Set rs = Nothing
End Sub


Public Sub PopClient()
Set rs = New ADODB.Recordset
rs.Open "Select ClientID, FName, MName, LName from Client", conn, 3, 3
Do While Not rs.EOF
    frmRent.cboClient.AddItem rs!FName '& Space(1) & rs!MName & Space(1) & rs!LName
rs.MoveNext
Loop


'If rs.EOF = True Then
'        frmRent.txtAddress.Text = rs!Address
'End If
End Sub

Public Sub Driver_Lock()
With frmDrivers
    .txtName.Enabled = False
    .txtAddress.Enabled = False
    .dtBirth.Enabled = False
    .txtPhone.Enabled = False
    .txtLicense.Enabled = False
    .dtReg.Enabled = False
    .cboCivStat.Enabled = False
    .txtTIN.Enabled = False
    .txtSSS.Enabled = False
    .cboPlate.Enabled = False
    .cboType.Enabled = False
    .cboMake.Enabled = False
    .cboColor.Enabled = False
End With
End Sub

Public Sub Driver_Unlock()
With frmDrivers
    .txtName.Enabled = True
    .txtAddress.Enabled = True
    .dtBirth.Enabled = True
    .txtPhone.Enabled = True
    .txtLicense.Enabled = True
    .dtReg.Enabled = True
    .cboCivStat.Enabled = True
    .txtTIN.Enabled = True
    .txtSSS.Enabled = True
    .cboPlate.Enabled = True
    .cboType.Enabled = True
    .cboMake.Enabled = True
    .cboColor.Enabled = True
    .txtName.SetFocus
End With
End Sub
Public Sub Driver_Clear()
With frmDrivers
    .txtName.Text = vbNullString
    .txtAddress.Text = vbNullString
    .dtBirth.Value = Date
    .txtPhone.Text = vbNullString
    .txtLicense.Text = vbNullString
    .dtReg.Value = Date
    .cboCivStat.Text = vbNullString
    .txtTIN.Text = vbNullString
    .txtSSS.Text = vbNullString
    .cboPlate.Text = vbNullString
    .cboType.Text = vbNullString
    .cboMake.Text = vbNullString
    .cboColor.Text = vbNullString

End With
End Sub

Public Sub PopBrand()
Set rs = New ADODB.Recordset
rs.Open "Select Brand from Make", conn, 3, 3
Do While Not rs.EOF
    frmDrivers.cboMake.AddItem rs!Brand
rs.MoveNext
Loop
End Sub

Public Sub PopDriveColor()
Set rs = New ADODB.Recordset
rs.Open "Select Trip from Trip", conn, 3, 3
Do While Not rs.EOF
    frmDrivers.cboColor.AddItem rs!Trip
rs.MoveNext
Loop
End Sub

Public Sub RefDriver()
    frmDrivers.lvDriver.ListItems.Clear
    Set rs = New ADODB.Recordset
    rs.Open "Select * from Driver", conn, adOpenForwardOnly, adLockPessimistic
    With rs
            Do While Not .EOF
                frmDrivers.lvDriver.ListItems.Add , , !DriverID, 1, 1
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(1) = "" & !DriverName
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(2) = "" & !Address
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(3) = "" & !BDate
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(4) = "" & !PhoneNo
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(5) = "" & !License
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(6) = "" & !DateReg
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(7) = "" & !CStatus
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(8) = "" & !TIN
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(9) = "" & !SSS
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(10) = "" & !PlateNo
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(11) = "" & !Type
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(12) = "" & !Make
                frmDrivers.lvDriver.ListItems(frmDrivers.lvDriver.ListItems.Count).SubItems(13) = "" & !Trip
            .MoveNext
            Loop
            .Close
    End With
Set rs = Nothing
End Sub

Public Sub PopCarInfo()
Set rs = New ADODB.Recordset
rs.Open "Select PlateNo, Type, Trip, Make from Car", conn, 3, 3
Do While Not rs.EOF
    frmDrivers.cboPlate.AddItem rs!PlateNo
    frmDrivers.cboType.AddItem rs!Type
rs.MoveNext
Loop
End Sub

Public Sub NoClientSearch()
Set rs = New ADODB.Recordset
rs.Open "Select * from Client", conn, adOpenForwardOnly, adLockPessimistic
frmSearch.Label15.Caption = rs.RecordCount
rs.Close
Set rs = Nothing
End Sub
Public Sub NoRentedSearch()
Set rs = New ADODB.Recordset
rs.Open "Select * from Rent_Car", conn, adOpenForwardOnly, adLockPessimistic
frmSearch.Label10.Caption = rs.RecordCount
rs.Close
Set rs = Nothing
End Sub

Public Sub NoCarsSearch()
Set rs = New ADODB.Recordset
rs.Open "Select * from Car", conn, adOpenForwardOnly, adLockPessimistic
frmSearch.Label4.Caption = rs.RecordCount
rs.Close
Set rs = Nothing
End Sub

Public Sub NoDriverSearch()
Set rs = New ADODB.Recordset
rs.Open "Select * from Driver", conn, adOpenForwardOnly, adLockPessimistic
frmSearch.Label21.Caption = rs.RecordCount
rs.Close
Set rs = Nothing

End Sub

Public Sub NoCarUnavail()
Set rs = New ADODB.Recordset
rs.Open "Select * from Car where AvailStatus='" & "Unavailable" & "'", conn, adOpenForwardOnly, adLockPessimistic
frmSearch.Label10.Caption = rs.RecordCount
rs.Close
End Sub

Public Sub PopDriverName()
Set rs = New ADODB.Recordset
rs.Open "Select DriverID, DriverName from Driver", conn, 3, 3
Do While Not rs.EOF
    frmRent.cboDriver.AddItem rs!DriverName
rs.MoveNext
Loop
End Sub
Public Sub PopCarPlate()
Set rs = New ADODB.Recordset
rs.Open "Select PlateNo from Driver", conn, 3, 3
Do While Not rs.EOF
    frmRent.cboPlate.AddItem rs!PlateNo
rs.MoveNext
Loop
End Sub
Public Sub Display_Plate()
Set rs = New ADODB.Recordset
rs.Open "Select DriverID, PlateNo from Driver where DriverName='" & frmRent.cboDriver.Text & "'", conn, 3, 3
With frmRent
    driver = rs!DriverID
    .cboPlate.Text = rs!PlateNo
End With

End Sub

Public Sub Display_Model()
Set rs = New ADODB.Recordset
rs.Open "Select CarID, Model, Condition, Make from Car where PlateNo='" & frmRent.cboPlate.Text & "'", conn, 3, 3
With frmRent
If rs.EOF Then
    Exit Sub
Else
    car = rs!CarID
    .txtModel.Text = rs!Model
    .txtCondition.Text = rs!Condition
    .txtMake.Text = rs!Make
End If
End With
End Sub

Public Sub Display_Rec()

Set rs = New ADODB.Recordset
rs.Open "Select Address, ClientID from Client where FName='" & frmRent.cboClient.Text & "'", conn, 3, 3
With frmRent
    .txtAddress.Text = rs!Address
    client = rs!ClientID
End With
End Sub
Public Sub Rent_Clear()
With frmRent
    .txtRateApplied.Text = vbNullString
    .lblSubTotal.Caption = "0.00"
    .lblTaxAmount.Caption = "0.00"
    .lblTotal.Caption = "0.00"
    .cboClient.Clear
    .cboDriver.Clear
    .txtAddress.Text = vbNullString
    .txtDay.Text = vbNullString
    .dtFrom.Value = Date
    .dtTo.Value = Date
    .cboPass.Clear
    .cboPlate.Text = vbNullString
    .txtModel.Text = vbNullString
    .txtCondition.Text = vbNullString
    .txtMake.Text = vbNullString
    .txtMileage.Text = vbNullString
    .txtTank.Text = vbNullString
End With
End Sub

Public Sub Rent_Lock()
With frmRent
    .txtRateApplied.Enabled = False
    .txtTaxRate.Enabled = False
    .cboClient.Enabled = False
    .cboDriver.Enabled = False
    .txtAddress.Enabled = False
    .txtDay.Enabled = False
    .dtFrom.Enabled = False
    .dtTo.Enabled = False
    .cboPass.Enabled = False
    .cboPlate.Enabled = False
    .txtModel.Enabled = False
    .txtCondition.Enabled = False
    .txtMake.Enabled = False
    .txtMileage.Enabled = False
    .txtTank.Enabled = False
End With

End Sub

Public Sub Rent_Unlock()
With frmRent
    .txtRateApplied.Enabled = True
    .txtTaxRate.Enabled = True
    .cboClient.Enabled = True
    .cboDriver.Enabled = True
    .txtAddress.Enabled = True
    .txtDay.Enabled = True
    .dtFrom.Enabled = True
    .dtTo.Enabled = True
    .cboPass.Enabled = True
    .cboPlate.Enabled = True
    .txtModel.Enabled = True
    .txtCondition.Enabled = True
    .txtMake.Enabled = True
    .txtMileage.Enabled = True
    .txtTank.Enabled = True
End With
End Sub

Public Sub RefTrans()
    frmRent.lvTrans.ListItems.Clear
    Set rs = New ADODB.Recordset
    rs.Open "Select * from Order_Transaction", conn, adOpenForwardOnly, adLockPessimistic
    With rs
            Do While Not .EOF
                frmRent.lvTrans.ListItems.Add , , !TransNo, 1, 1
                frmRent.lvTrans.ListItems(frmRent.lvTrans.ListItems.Count).SubItems(1) = "" & !ReceiptNo
                frmRent.lvTrans.ListItems(frmRent.lvTrans.ListItems.Count).SubItems(2) = "" & Format(!RateApplied, "#,###,##0.00")
                frmRent.lvTrans.ListItems(frmRent.lvTrans.ListItems.Count).SubItems(3) = "" & !TaxRate
                frmRent.lvTrans.ListItems(frmRent.lvTrans.ListItems.Count).SubItems(4) = "" & Format(!TaxAmount, "#,###,##0.00")
                frmRent.lvTrans.ListItems(frmRent.lvTrans.ListItems.Count).SubItems(5) = "" & Format(!Subtotal, "#,###,##0.00")
                frmRent.lvTrans.ListItems(frmRent.lvTrans.ListItems.Count).SubItems(6) = "" & Format(!Total, "#,###,##0.00")
            .MoveNext
            Loop
            .Close
    End With
Set rs = Nothing
End Sub
