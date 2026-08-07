Option Explicit

'/*
'Splits a string by delimiter and removes empty elements
'and preserves whitespace in non-empty elements
'
'return: Allocated String array
'*/
Public Function splitNoEmpty(ByVal text As String, ByVal delimiter As String) As String()
    If VBA.LenB(text) = 0 Then
        splitNoEmpty = VBA.Split(VBA.vbNullString)
        Exit Function
    End If

    Dim parts() As String: parts = VBA.Split(text, delimiter)
    Dim result() As String: ReDim result(0 To VBA.UBound(parts))
    
    Dim i As Long
    Dim j As Long: j = -1
    
    For i = 0 To VBA.UBound(parts)
        If VBA.LenB(parts(i)) > 0 Then
            result(j) = parts(i)
            j = j + 1
        End If
    Next i
    
    If j < 0 Then
        splitNoEmpty = VBA.Split(VBA.vbNullString)
    Else
        ReDim Preserve result(0 To j)
        splitNoEmpty = result
    End If
End Function

'/*
'Splits a string by delimiter, trims whitespace from each
'element, and removes blank elements
'
'return: Allocated String array
'*/
Public Function splitNoBlank(ByVal text As String, ByVal delimiter As String) As String()
    If VBA.LenB(text) = 0 Then
        splitNoBlank = VBA.Split(VBA.vbNullString)
        Exit Function
    End If

    Dim parts() As String: parts = VBA.Split(text, delimiter)
    Dim result() As String: ReDim result(0 To VBA.UBound(parts))
    
    Dim item As String
    Dim i As Long
    Dim j As Long: j = -1
    
    For i = 0 To VBA.UBound(parts)
        item = VBA.Trim$(parts(i))
        If VBA.LenB(item) > 0 Then
            result(j) = item
            j = j + 1
        End If
    Next i
    
    If j < 0 Then
        splitNoBlank = VBA.Split(VBA.vbNullString)
    Else
        ReDim Preserve result(0 To j)
        splitNoBlank = result
    End If
End Function