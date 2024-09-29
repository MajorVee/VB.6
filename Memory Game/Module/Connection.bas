Attribute VB_Name = "Connection"
Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset

Public Sub Main()
On Error GoTo Err
    CN.Provider = "Microsoft.ACE.OLEDB.12.0"
    CN.Open App.Path & "\Data.accdb"
   frmMain.Show
    Exit Sub
Err:
    MsgBox Err.Description, vbCritical
End Sub


