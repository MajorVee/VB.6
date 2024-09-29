VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mrs As New ADODB.Recordset
Dim msql As String
Public Sub AddCustomer(cID As String, cName As String, cGender As String, cAge As String, cContact As String, cAddress As String, cType As String)
If rs.State = adStateOpen Then rs.Close
msql = " Select * from tblCustomer Where cName='" & cName & "'"
rs.Open msql, conn
If rs.RecordCount >= 1 Then
MsgBox "Duplicate record found", vbCritical, ""
Exit Sub
End If
With rs
    .AddNew
    !cID = cID
    !cName = cName
    !cGender = cGender
    !cAge = cAge
    !cContact = cContact
    !cAddress = cAddress
    !cType = cType
    .Update
End With
MsgBox "New entry successfully saved to the record.", vbInformation
End Sub
Public Sub UpdateCustomer(cID As String, cName As String, cGender As String, cAge As String, cContact As String, cAddress As String, cType As String, OldID As String)
If rs.State = adStateOpen Then rs.Close
msql = " Update tblCustomer set cID='" & cID & "',cName='" & cName & "',cGender='" & cGender & "',cAge='" & cAge & "',cContact='" & cContact & "',cAddress='" & cAddress & "',ctype='" & cType & "' Where cID='" & OldID & "'"
rs.Open msql, conn
MsgBox "Record(s) Updated", vbInformation, ""
End Sub
Public Sub DeleteCustomer(cID As String)
If rs.State = adStateOpen Then rs.Close
msql = " Delete * From tblCustomer Where cID='" & cID & "'"
rs.Open msql, conn
MsgBox "Record(s) Deleted", vbInformation, ""
End Sub
Public Sub DisplayCustomer(lstDay As ListView, CustomerSearch As String)
Dim lstItem As ListItem, a As Integer
If rs.State = adStateOpen Then rs.Close
    If CustomerSearch <> "" Then
    msql = " SELECT tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType" & _
           " From tblCustomer" & _
           " Where tblCustomer.cName like '" & CustomerSearch & "%'" & _
    " GROUP BY tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType" & _
    " ORDER BY tblCustomer.cID;"
    Else
     msql = " SELECT tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType" & _
           " From tblCustomer" & _
    " GROUP BY tblCustomer.cID, tblCustomer.cName,tblCustomer.cGender,tblCustomer.cAge,tblCustomer.cContact,tblCustomer.cAddress,tblCustomer.cType" & _
    " ORDER BY tblCustomer.cID;"
    End If
 rs.Open msql, conn
   lstDay.ListItems.Clear
   Do While Not rs.EOF
    a = a + 1
        Set lstItem = lstDay.ListItems.Add(, , a, 1, 1)
            lstItem.SubItems(1) = rs(0).Value
            lstItem.SubItems(2) = rs(1).Value
            lstItem.SubItems(3) = rs(2).Value
            lstItem.SubItems(4) = rs(3).Value
            lstItem.SubItems(5) = rs(4).Value
            lstItem.SubItems(6) = rs(5).Value
            lstItem.SubItems(7) = rs(6).Value
            rs.MoveNext
            Loop
End Sub
Function GetID() As String
If rs.State = adStateOpen Then rs.Close
msql = " SELECT Max(tblCustomer.cID) AS MaxOfcID" & _
" FROM tblCustomer;"
rs.Open msql, conn
If IsNull(rs(0).Value) Then
GetID = Format(1, "00000")
Else
GetID = Format(rs(0).Value + 1, "00000")
End If
End Function







