Option Explicit

'/*
'A value bounded by a minimum and a maximum
'
'return: If value is less than min then returns min
'        If value is greater than max then returns max
'        Otherwise this returns value
'*/
Public Function clamp(ByVal value As Long, ByVal min As Long, ByVal max As Long) As Long
    If min > max Then
        Err.Raise _
            Number:=5, _
            Source:="uLong.clamp", _
            Description:="Invalid boundaries: 'min' cannot be greater than 'max'."
    End If

    If value < min Then
        clamp = min
    ElseIf value > max Then
        clamp = max
    Else
        clamp = value
    End If
End Function

'/*
'A value bounded by a minimum value
'
'return: If value is less than bound then returns bound
'        Otherwise this returns value
'*/
Public Function minAt(ByVal value As Long, ByVal bound As Long) As Long
    If value < bound Then
        minAt = bound
    Else
        minAt = value
    End If
End Function

'/*
'A value bounded by a maximum value
'
'return: If value is greater than bound then returns bound
'        Otherwise this returns value
'*/
Public Function maxAt(ByVal value As Long, ByVal bound As Long) As Long
    If value > bound Then
        maxAt = bound
    Else
        maxAt = value
    End If
End Function

'/*
'The minimum of two values
'
'return: The lesser of a and b
'*/
Public Function minOf(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then
        minOf = a
    Else
        minOf = b
    End If
End Function

'/*
'The maximum of two values
'
'return: The greater of a and b
'*/
Public Function maxOf(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then
        maxOf = a
    Else
        maxOf = b
    End If
End Function

'/*
'The minimum value in an array
'
'return: The lowest number stored in the array
'        Zero if unallocated
'*/
Public Function minIn(ByRef nums() As Long) As Long
    If uArray.count(nums) = 0 Then Exit Function
    Dim lower As Long: lower = LBound(nums)
    Dim i As Long
    minIn = nums(lower)
    For i = lower + 1 To UBound(nums)
        If nums(i) < minIn Then minIn = nums(i)
    Next i
End Function

'/*
'The maximum value in an array
'
'return: The highest number stored in the array
'        Zero if unallocated
'*/
Public Function maxIn(ByRef nums() As Long) As Long
    If uArray.count(nums) = 0 Then Exit Function
    Dim lower As Long: lower = LBound(nums)
    Dim i As Long
    maxIn = nums(lower)
    For i = lower + 1 To UBound(nums)
        If nums(i) > maxIn Then maxIn = nums(i)
    Next i
End Function