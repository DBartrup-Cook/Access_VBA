Attribute VB_Name = "mdl_CreateXL"
Option Compare Database
Option Explicit

'----------------------------------------------------------------------------------
' Procedure : CreateXL
' Author    : Darren Bartrup-Cook
' Date      : 02/10/2014
' Purpose   : Creates an instance of Excel and passes the reference back.
'-----------------------------------------------------------------------------------
Public Function CreateXL(Optional bVisible As Boolean = True) As Object

    Dim oTmpXL As Object

    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Defer error trapping in case Excel is not running. '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    Set oTmpXL = GetObject(, "Excel.Application")
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'If an error occurs then create an instance of Excel. '
    'Reinstate error handling.                            '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ERROR_HANDLER
        Set oTmpXL = CreateObject("Excel.Application")
    End If
    
    oTmpXL.Visible = bVisible
    CreateXL = oTmpXL

    On Error GoTo 0
    Exit Function

ERROR_HANDLER:
    Select Case Err.Number
        
        Case Else
            MsgBox "Error " & Err.Number & vbCr & _
                " (" & Err.Description & ") in procedure CreateXL."
            Err.Clear
    End Select
    

End Function
