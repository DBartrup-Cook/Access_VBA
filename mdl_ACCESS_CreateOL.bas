Attribute VB_Name = "mdl_CreateOL"
Option Compare Database
Option Explicit

'----------------------------------------------------------------------------------
' Procedure : CreateOL
' Author    : Darren Bartrup-Cook
' Date      : 13/01/2015
' Purpose   : Creates an instance of Outlook and passes the reference back.
'-----------------------------------------------------------------------------------
Public Function CreateOL() As Object

    Dim oTmpOL As Object
    
    On Error GoTo ERROR_HANDLER

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Creating an instance of Outlook is different from Excel. '
    'There can only be a single instance of Outlook running,  '
    'so CreateObject will GetObject if it already exists.     '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set oTmpOL = CreateObject("Outlook.Application")
    
    Set CreateOL = oTmpOL

    On Error GoTo 0
    Exit Function

ERROR_HANDLER:
    Select Case Err.Number
        
        Case Else
            MsgBox "Error " & Err.Number & vbCr & _
                " (" & Err.Description & ") in procedure CreateOL."
            Err.Clear
    End Select
    
End Function

