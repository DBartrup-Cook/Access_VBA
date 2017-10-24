Attribute VB_Name = "mdl_FindLastCell"
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : LastCell
' Author    : Darren Bartrup-Cook
' Date      : 26/11/2013
' Purpose   : Finds the last cell containing data or a formula within the given worksheet.
'             If the Optional Col is passed it finds the last row for a specific column.
'---------------------------------------------------------------------------------------
Public Function LastCell(wrkSht As Object, Optional Col As Long = 0) As Object

    Dim lLastCol As Long, lLastRow As Long
    
    On Error Resume Next
    
    With wrkSht
        If Col = 0 Then
            lLastCol = .Cells.Find("*", , , , 2, 2).Column
            lLastRow = .Cells.Find("*", , , , 1, 2).Row
        Else
            lLastCol = .Cells.Find("*", , , , 2, 2).Column
            lLastRow = .Columns(Col).Find("*", , , , 2, 2).Row
        End If
        
        If lLastCol = 0 Then lLastCol = 1
        If lLastRow = 0 Then lLastRow = 1
        
        Set LastCell = wrkSht.Cells(lLastRow, lLastCol)
    End With
    On Error GoTo 0
    
End Function
