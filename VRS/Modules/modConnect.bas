Attribute VB_Name = "modConnect"
Option Explicit
Global conn As ADODB.Connection
Global rs As ADODB.Recordset
Global rs2 As ADODB.Recordset
Global rs3 As ADODB.Recordset
Global rs4 As ADODB.Recordset
Public xcl As Excel.Application

Public Sub Connected()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset

conn.CursorLocation = adUseClient
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Data.mdb;Persist Security Info=False"
conn.Open

End Sub



