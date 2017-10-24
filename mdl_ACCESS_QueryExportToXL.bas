Attribute VB_Name = "mdl_QueryExportToXL"
Option Compare Database
Option Explicit

'----------------------------------------------------------------------------------
' Procedure : QueryExportToXL
' Author    : Darren Bartrup-Cook
' Date      : 26/08/2014
' Purpose   : Exports a named query or recordset to Excel.
'-----------------------------------------------------------------------------------
Public Function QueryExportToXL(wrkSht As Object, Optional sQueryName As String, _
                                                  Optional rst As DAO.Recordset, _
                                                  Optional SheetName As String, _
                                                  Optional rStartCell As Object, _
                                                  Optional AutoFitCols As Boolean = True, _
                                                  Optional colHeadings As Collection) As Boolean

    Dim db As DAO.Database
    Dim prm As DAO.Parameter
    Dim qdf As DAO.QueryDef
    Dim fld As DAO.Field
    Dim oXLCell As Object
    Dim vHeading As Variant
    
    On Error GoTo ERROR_HANDLER
    
    If sQueryName <> "" And rst Is Nothing Then
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Open the query recordset.                               '
        'Any parameters in the query need to be evaluated first. '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set db = CurrentDb
        Set qdf = db.QueryDefs(sQueryName)
        For Each prm In qdf.Parameters
            prm.Value = Eval(prm.Name)
        Next prm
        Set rst = qdf.OpenRecordset
    End If
    
    If rStartCell Is Nothing Then
        Set rStartCell = wrkSht.cells(1, 1)
    Else
        If rStartCell.Parent.Name <> wrkSht.Name Then
            Err.Raise 4000, , "Incorrect Start Cell parent."
        End If
    End If
    
    
    If Not rst.BOF And Not rst.EOF Then
        With wrkSht
            Set oXLCell = rStartCell
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Paste the field names from the query into row 1 of the sheet. '
            'TO DO: Facility to use an alternative name.                   '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If colHeadings Is Nothing Then
                For Each fld In rst.Fields
                    oXLCell.Value = fld.Name
                    Set oXLCell = oXLCell.Offset(, 1)
                Next fld
            Else
                For Each vHeading In colHeadings
                    oXLCell.Value = vHeading
                    Set oXLCell = oXLCell.Offset(, 1)
                Next vHeading
            End If
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Paste the records from the query into row 2 of the sheet. '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Set oXLCell = rStartCell.Offset(1, 0)
            oXLCell.copyfromrecordset rst
            If AutoFitCols Then
                .Columns.Autofit
            End If
            
            If SheetName <> "" Then
                .Name = SheetName
            End If
            
            '''''''''''''''''''''''''''''''''''''''''''
            'TO DO: Has recordset imported correctly? '
            '''''''''''''''''''''''''''''''''''''''''''
            QueryExportToXL = True
            
        End With
    Else
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'There are no records to export, so the export has failed. '
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        QueryExportToXL = False
    End If
    
    Set db = Nothing

    On Error GoTo 0
    Exit Function

ERROR_HANDLER:
    Select Case Err.Number
        
        Case Else
            MsgBox "Error " & Err.Number & vbCr & _
                " (" & Err.Description & ") in procedure QueryExportToXL."
            Err.Clear
            Resume
    End Select
    
End Function


