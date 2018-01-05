Attribute VB_Name = "mdl_SQLDate"
Option Compare Database
Option Explicit

    'Purpose:    Return a delimited string in the date format used natively by JET SQL.
    'Argument:   A date/time value.
    'Note:       Returns just the date format if the argument has no time component,
    '                or a date/time format if it does.
    'Author:     Allen Browne. allen@allenbrowne.com, June 2006.
Function SQLDate(varDate As Variant) As String
    If IsDate(varDate) Then
        If DateValue(varDate) = varDate Then
            SQLDate = Format$(varDate, "\#mm\/dd\/yyyy\#")
        Else
            SQLDate = Format$(varDate, "\#mm\/dd\/yyyy hh\:nn\:ss\#")
        End If
    End If
End Function
