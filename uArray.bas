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