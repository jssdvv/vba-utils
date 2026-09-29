Option Explicit

'/*
' Splits a [text] by [delimiter], preserves whitespaces in non-empty
' elements and removes empty elements.
'
' @return Allocated String array.
'*/
Public Function splitNoEmpty(ByVal text As String, ByVal delimiter As String) As String()
    If VBA.LenB(text) = 0 Then
        splitNoEmpty = VBA.Split(VBA.vbNullString)
        Exit Function
    End If

    Dim parts() As String: parts = VBA.Split(text, delimiter)
    Dim result() As String: ReDim result(0 To UBound(parts))
    
    Dim i As Long
    Dim j As Long: j = -1
    
    For i = 0 To UBound(parts)
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
' Splits a [text] by [delimiter], trims whitespaces from each element
' and removes empty elements.
'
' @return Allocated String array.
'*/
Public Function splitNoBlank(ByVal text As String, ByVal delimiter As String) As String()
    If VBA.LenB(text) = 0 Then
        splitNoBlank = VBA.Split(VBA.vbNullString)
        Exit Function
    End If

    Dim parts() As String: parts = VBA.Split(text, delimiter)
    Dim result() As String: ReDim result(0 To UBound(parts))
    
    Dim item As String
    Dim i As Long
    Dim j As Long: j = -1
    
    For i = 0 To UBound(parts)
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
' Counts the occurrences of specific [chars] inside a [text].
'
' @return Number of chars occurrences, or zero if [text] or [chars] is
'         empty or if [chars] is longer than [text].
'*/
Public Function countChars(ByVal text As String, ByVal chars As String) As Long
    Dim lenBText As Long: lenBText = VBA.LenB(text)
    Dim lenBChars As Long: lenBChars = VBA.LenB(chars)

    If lenBText = 0 Or lenBChars = 0 Or lenBChars > lenBText Then Exit Function

    Dim posB As Long: posB = 1

    Do
        posB = VBA.InStrB(posB, text, chars, vbBinaryCompare)
        If posB = 0 Then Exit Do
        countChars = countChars + 1
        posB = posB + lenBChars
    Loop
End Function

'/*
' Gets the digits inside a [text].
'
' @return The numeric digits of a [text] as a String.
'*/
Public Function digits(ByVal text As String) As String
    If VBA.LenB(text) = 0 Then Exit Function
    
    Dim bytes() As Byte: bytes = text
    Dim result() As Byte
    ReDim result(0 To UBound(bytes))
    
    Dim i As Long, j As Long
    For i = LBound(bytes) To UBound(bytes) Step 2
        If (bytes(i) >= 48) And (bytes(i) <= 57) Then
            result(j) = bytes(i)
            j = j + 2
        End If
    Next i
    
    If j > 0 Then
        ReDim Preserve result(0 To j - 1)
        digits = result
    End If
End Function

'/*
' Converts an array of bytes into a String.
'
' @return The bytes array as a String.
'*/
Public Function bytesToHexString(ByRef arr() As Byte) As String
    If uArray.count(arr) = 0 Then
        VBA.Err.Raise _
            Number:=9, _
            Source:="uString.bytesToHexString", _
            Description:="Invalid argument: 'arr' is not initialized or is empty."
    End If

    Dim i As Long
    Dim s As String
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) < 16 Then
            s = s & "0" & VBA.Hex$(arr(i))
        Else
            s = s & VBA.Hex$(arr(i))
        End If
    Next i
    
    bytesToHexString = s
End Function