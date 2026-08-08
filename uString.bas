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

'/*
'Counts how many times a specific substring or character
'occurs inside a string
'
'return: Number of chars occurrences
'        Zero if text or chars is empty
'        Zero if chars is longer than text
'*/
Public Function countChars(ByVal text As String, ByVal chars As String) As Long
    Dim lenBText As Long: lenBText = VBA.LenB(text)
    Dim lenBChars As Long: lenBChars = VBA.LenB(chars)

    If lenBText = 0 Or lenBChars = 0 Or lenBChars > lenBText Then Exit Function

    Dim lenChars As Long: lenChars = lenBChars \ 2
    Dim pos As Long: pos = 1

    Do
        pos = VBA.InStr(pos, text, chars, vbBinaryCompare)
        If pos = 0 Then Exit Do
        countChars = countChars + 1
        pos = pos + lenChars
    Loop
End Function