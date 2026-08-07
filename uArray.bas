Option Explicit

'/*
'The element count of an array dimension
'
'return: Number of elements in the specified dimension
'        Zero if unallocated or invalid
'*/
Public Function count(ByRef arr As Variant, Optional ByVal dimension As Long = 1) As Long
    If Not IsArray(arr) Then Exit Function
    On Error Resume Next
    Dim upper As Long: upper = UBound(arr, dimension)
    If Err.Number = 0 Then count = upper - LBound(arr, dimension) + 1
    On Error GoTo 0
End Function