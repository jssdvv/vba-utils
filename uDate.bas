Option Explicit

'/*
' Shifts a date by offseting [years], [months], and [days].
'
' @return The resulting date after applying the specified offsets.
'*/
Public Function shiftDate( _
    ByVal value As Date, _
    Optional ByVal years As Long = 0, _
    Optional ByVal months As Long = 0, _
    Optional ByVal days As Long = 0) As Date

    If years <> 0 Then value = VBA.DateAdd("yyyy", years, value)
    If months <> 0 Then value = VBA.DateAdd("m", months, value)
    If days <> 0 Then value = VBA.DateAdd("d", days, value)
    shiftDate = value
End Function

'/*
' Adds or subtracts [years] from a date.
'
' @return The modified date, or [value] if zero.
'*/
Public Function addYears(ByVal value As Date, ByVal years As Long) As Date
    If years <> 0 Then
        addYears = VBA.DateAdd("yyyy", years, value)
    Else
        addYears = value
    End If
End Function

'/*
' Adds or subtracts [months] from a date.
'
' @return The modified date, or [value] if zero.
'*/
Public Function addMonths(ByVal value As Date, ByVal months As Long) As Date
    If months <> 0 Then
        addMonths = VBA.DateAdd("m", months, value)
    Else
        addMonths = value
    End If
End Function

'/*
' Adds or subtracts [days] from a date.
'
' @return The modified date, or [value] if zero.
'*/
Public Function addDays(ByVal value As Date, ByVal days As Long) As Date
    If days <> 0 Then
        addDays = VBA.DateAdd("d", days, value)
    Else
        addDays = value
    End If
End Function