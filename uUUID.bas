Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef id As GUIDBytes) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef id As GUIDBytes) As Long
#End If

'Standard: https://doi.org/10.17487/RFC9562
'
'xxxxxxxx-xxxx-Mxxx-Nxxx-xxxxxxxxxxxx
'              ||   ||
'         Version   Variant

Private Type GUIDBytes
    data1(0 To 3) As Byte '32 bits
    data2(0 To 1) As Byte '16 bits
    data3(0 To 1) As Byte '16 bits
    data4(0 To 7) As Byte '64 bits
End Type

Private Enum UUIDVersion
    V1 = &H10   'The Gregorian time-based UUID
    V2 = &H20   'DCE Security with embedded POSIX UUIDs
    V3 = &H30   'The name-based (MD5)
    V4 = &H40   'The randomly or pseudorandomly generated
    V5 = &H50   'The name-based (SHA-1)
    V6 = &H60   'Reordered Gregorian time-based UUID
    V7 = &H70   'Unix Epoch time-based UUID
    V8 = &H80   'Reserved for custom UUID
End Enum

Private Enum UUIDVariant
    V0XXX = &H0    'Network Computing System (NCS)
    V10XX = &H80   'RFC 4122 / RFC 9562
    V110X = &HC0   'Microsoft Corporation backward compatibility
    V111X = &HE0   'Reserved for future definition
End Enum

'/*
' Sets the version nibble on the most significant byte of the
' (already reversed) Data3 field, following RFC 9562.
'
' @param [b]     The byte to modify, passed by reference.
' @param [value] The UUID version to encode (see [UUIDVersion]).
'*/
Private Sub setVersion(ByRef b As Byte, ByVal value As UUIDVersion)
    'Bits 0F: 0000 1111
    b = (b And &HF) Or value
End Sub

'/*
' Sets the fixed variant bits on the most significant byte of the
' Data4 field, following RFC 9562.
'
' @param [b]     The byte to modify, passed by reference.
' @param [value] The UUID variant to encode.
'*/
Private Sub setVariant(ByRef b As Byte, ByVal value As UUIDVariant)
    Dim mask As Byte

    Select Case value
        Case V0XXX, V10XX
            'Bits 3F: 0011 1111
            mask = &H3F
            
        Case V110X, V111X
            'Bits 1F: 0001 1111
            mask = &H1F
    End Select

    b = (b And mask) Or value
End Sub

'/*
' Generates a random UUID version 4 (RFC 9562), using the operating
' system's cryptographic random source via CoCreateGuid.
'
' @param [lowerCase] Whether to return the UUID in lowercase.
'                    Defaults to False (uppercase, as returned by Hex$).
'
' @return A 36-character UUID string in canonical form
'         (xxxxxxxx-xxxx-4xxx-Nxxx-xxxxxxxxxxxx).
'
' @throws Runtime error vbObjectError + 1000 if CoCreateGuid fails.
'*/
Public Function UUIDv4(Optional ByVal lowerCase As Boolean = False) As String
    Dim g As GUIDBytes
    Dim s As String
    
    If CoCreateGuid(g) <> 0 Then
        VBA.Err.Raise _
            Number:=vbObjectError + 1000, _
            Source:="mUUID.UUIDv4", _
            Description:="CoCreateGuid failed."
    End If
    
    'Little-Endian To Big-Endian
    uArray.reverseBytes g.data1
    uArray.reverseBytes g.data2
    uArray.reverseBytes g.data3
    
    setVersion g.data3(0), V4
    setVariant g.data4(0), V10XX

    s = uString.bytesToHexString(g.data1) & "-" & _
        uString.bytesToHexString(g.data2) & "-" & _
        uString.bytesToHexString(g.data3) & "-" & _
        uString.bytesToHexString(g.data4)

    If lowerCase Then
        UUIDv4 = VBA.LCase$(s)
    Else
        UUIDv4 = s
    End If
End Function