Attribute VB_Name = "modBackup"
Option Explicit

Public Function DoesFileExist(PathName As String) As Boolean

If Dir$(PathName) <> vbNullString Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
End Function

Public Function MDbackupdatabases() As Long
Dim strPath, strBackup As String
Dim sbakfile As String
Dim FSO As FileSystemObject
Set FSO = CreateObject("Scripting.FileSystemObject")

If DoesFileExist(App.Path & "\Backup\Data-" & Format(Date, "dd-mm-yyyy") & ".mdb") Then
MsgBox "The backup database already exist", vbCritical, "Database Exist!"
frmBackup.Timer2.Enabled = False
Exit Function
Else
strPath = App.Path & "\Data.mdb"
strBackup = App.Path & "\Backup\Data-" & Format(Date, "dd-mm-yyyy") & ".mdb"
FSO.CopyFile strPath, strBackup
frmBackup.Timer2.Enabled = True
End If
End Function


Public Sub ColorListviewRow(lv As ListView, RowNbr As Long, RowColor As OLE_COLOR)

    
    Dim itmX As ListItem
    Dim lvSI As ListSubItem
    Dim intIndex As Integer
    
    On Error GoTo ErrorRoutine
    
    Set itmX = lv.ListItems(RowNbr)
    itmX.ForeColor = RowColor
    For intIndex = 1 To lv.ColumnHeaders.Count - 1
        Set lvSI = itmX.ListSubItems(intIndex)
        lvSI.ForeColor = RowColor
    Next
 
    Set itmX = Nothing
    Set lvSI = Nothing
    
    Exit Sub
 
ErrorRoutine:
 
    MsgBox Err.Description
 
End Sub


