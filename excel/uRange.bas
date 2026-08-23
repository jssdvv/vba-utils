Option Explicit

Public Function columnLetter(ByVal index As Long) As String
    If index < 1 Or index > 16384 Then
        VBA.Err.Raise _
            Number:=5, _
            Source:="uRange.columnLetter", _
            Description:="Invalid boundaries: 'index' cannot be less than '1' or greater than '16384'."
    End If

    Dim remainder As Long
    Do While index > 0
        index = index - 1
        remainder = index Mod 26
        columnLetter = VBA.Chr(65 + remainder) & columnLetter
        index = index \ 26
    Loop
End Function