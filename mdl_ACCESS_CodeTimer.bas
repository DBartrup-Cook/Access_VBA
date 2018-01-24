Attribute VB_Name = "mdl_CodeTimer"
Option Compare Database
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public CodeTimer As Long

Public Function StartTimer()
    CodeTimer = GetTickCount
End Function

Public Function StopTimer()
    Dim FinalTime As Long
    FinalTime = GetTickCount - CodeTimer
    MsgBox Format(Now(), "ddd dd-mmm-yy hh:mm:ss") & vbCr & vbCr & _
            Format((FinalTime / 1000) / 86400, "hh:mm:ss") & vbCr & _
            FinalTime & " ms.", vbOKOnly + vbInformation, _
        "Code Timer"
    CodeTimer = 0
End Function

Public Sub Test()

    StartTimer
    
    'Other code
    
    StopTimer
    
End Sub
