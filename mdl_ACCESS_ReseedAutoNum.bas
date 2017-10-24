Attribute VB_Name = "mdl_ReseedAutoNum"
Option Compare Database


Function DeleteAllAndResetAutoNum(strTable As String) As Boolean
    'Purpose:   Delete all records from the table, and reset the AutoNumber using ADOX.
    '           Also illustrates how to find the AutoNumber field.
    'Argument:  Name of the table to reset.
    'Return:    True if sucessful.
    
    Dim cat As Object 'New ADOX.Catalog
    Dim tbl As Object 'ADOX.Table
    Dim col As Object 'ADOX.Column
    Dim strSql As String
    
    Set cat = CreateObject("ADOX.Catalog")
    Set tbl = CreateObject("ADOX.Table")
    Set col = CreateObject("ADOX.Column")
    
    'Delete all records.
    strSql = "DELETE FROM [" & strTable & "];"
    CurrentProject.Connection.Execute strSql
    
    'Find and reset the AutoNum field.
    cat.ActiveConnection = CurrentProject.Connection
    Set tbl = cat.Tables(strTable)
    For Each col In tbl.Columns
        If col.Properties("Autoincrement") Then
            col.Properties("Seed") = 1
            DeleteAllAndResetAutoNum = True
        End If
    Next
End Function
