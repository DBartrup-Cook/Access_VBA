Attribute VB_Name = "mdl_FormIcon"
Option Compare Database
Option Explicit

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const WM_SETICON = &H80
Private Const IMAGE_ICON = 1
Private Const LR_LOADFROMFILE = &H10
Private Const SM_CXSMICON As Long = 49
Private Const SM_CYSMICON As Long = 50

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long

'    Purpose:    Displays a custom icon on the referenced form (hWnd).
'    Note:       I've no idea who wrote this code.  The earliest I've found is
'                20-Feb-2007 at https://access-programmers.co.uk/forums/showthread.php?t=123449
'    Author:     ??
Public Function SetFormIcon(hWnd As Long, strIconPath As String) As Boolean
    Dim lIcon As Long
    Dim lResult As Long
    Dim X As Long, Y As Long
    
    X = GetSystemMetrics(SM_CXSMICON)
    Y = GetSystemMetrics(SM_CYSMICON)
    lIcon = LoadImage(0, strIconPath, 1, X, Y, LR_LOADFROMFILE)
    lResult = SendMessage(hWnd, WM_SETICON, 0, ByVal lIcon)
End Function

'Example Use.
'Private Sub Form_Load()
'
'On Error GoTo ERR_HANDLE
'
'    SetFormIcon Me.hWnd, Left(CurrentDb.Name, Len(CurrentDb.Name) - Len(Dir(CurrentDb.Name))) & "\icons\calendar.ico"
'
'EXIT_PROC:
'    On Error GoTo 0
'    Exit Sub
'
'ERR_HANDLE:
'    Select Case Err.Number
'
'        Case Else
'            DisplayError Err.Number, Err.Description, "TestModule.Form_Load()"
'            Resume EXIT_PROC
'    End Select
'
'End Sub


