Attribute VB_Name = "mdl_ReseedAutoNum"
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure   : DeleteAllAndResetAutoNum
' Provided by : Allen Browne (http://allenbrowne.com/func-ADOX.html)
' Date        : March 2007, July 2009
' Purpose     : Delete all records from the table, and reset the AutoNumber using ADOX.
'               Also illustrates how to find the AutoNumber field.
'---------------------------------------------------------------------------------------
Public Sub DeleteAllAndResetAutoNum(strTable As String, Reseed As Boolean)
    
    Dim cat As Object 'New ADOX.Catalog
    Dim tbl As Object 'ADOX.Table
    Dim col As Object 'ADOX.Column
    Dim strSql As String
    
    On Error GoTo ERR_HANDLE
    
    Set cat = CreateObject("ADOX.Catalog")
    Set tbl = CreateObject("ADOX.Table")
    Set col = CreateObject("ADOX.Column")
    
    'Delete all records.
    strSql = "DELETE FROM [" & strTable & "];"
    CurrentProject.Connection.Execute strSql
    
    'Find and reset the AutoNum field.
    If Reseed Then
        cat.ActiveConnection = CurrentProject.Connection
        Set tbl = cat.Tables(strTable)
        For Each col In tbl.Columns
            If col.Properties("Autoincrement") Then
                col.Properties("Seed") = 1
            End If
        Next
    End If
    
EXIT_PROC:
    On Error GoTo 0
    Exit Sub
ERR_HANDLE:
    Select Case Err.Number
    
        Case Else
            DisplayError Err.Number, Err.Description, "mdlReseedAutoNum.DeleteAllAndResetAutoNum()"
            Resume EXIT_PROC
    End Select
    
End Sub
