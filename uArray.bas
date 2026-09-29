Option Explicit

'/*
' The element count of an array [dimension].
'
' @return Number of elements in the specified [dimension] or zero
'         if the [arr] is unallocated or invalid.
'*/
Public Function count(ByRef arr As Variant, Optional ByVal dimension As Long = 1) As Long
    If Not VBA.IsArray(arr) Then Exit Function
    On Error Resume Next
    count = UBound(arr, dimension) - LBound(arr, dimension) + 1
    On Error GoTo 0
End Function

'/*
' Reverses 1D bytes array by reference.
'*/
Public Function reverseBytes(ByRef arr() As Byte)
    Dim c As Long: c = uArray.count(arr)
    
    If c = 0 Then
        VBA.Err.Raise _
            Number:=9, _
            Source:="uArray.reverseBytes", _
            Description:="Invalid argument: 'arr' is not initialized or is empty."
    End If

    Dim i As Long
    Dim b As Byte
    Dim lower As Long: lower = LBound(arr)
    Dim upper As Long: upper = UBound(arr)
    
    For i = lower To lower + (c \ 2) - 1
        b = arr(i)
        arr(i) = arr(upper)
        arr(upper) = b
        upper = upper - 1
    Next i
End Function