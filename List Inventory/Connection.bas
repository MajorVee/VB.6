Attribute VB_Name = "Connection"
Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset

Public xcl As Excel.Application


Public Sub Main()
On Error GoTo Err

CN.Provider = "MICROSOFT.ACE.OLEDB.12.0"
CN.Open App.Path & "\Sample.accdb"
frmInformation.Show
Exit Sub
Err:
    MsgBox Err.Description, vbCritical
End Sub

