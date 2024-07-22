Option Explicit

Private Const Base32Chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"

Sub Test()
    Dim arrB() As Byte, s As String, n As Long, i As Long
    n = 17
    Redim arrB(n)
    For i = 0 to n: arrB(i) = CByte(Rnd() * 255): Debug.Print arrB(i); : Next: Debug.Print vbNullString
    s = ToBase32(arrB): Debug.Print s
End Sub

Public Function ToBase32(inputByte() As Byte) As String
    Dim sb As String, offset As Long, numCharsToOutput As Integer: sb = vbNullString: offset = LBound(inputByte)
    Dim a As Byte, b As Byte, c As Byte, d As Byte, e As Byte, f As Byte, g As Byte, h As Byte
    Do While offset <= UBound(inputByte)
        numCharsToOutput = GetNextGroup(inputByte, offset, a, b, c, d, e, f, g, h)
        sb = sb & IIf(numCharsToOutput >= 1, Mid$(Base32Chars, a + 1, 1), "=")
        sb = sb & IIf(numCharsToOutput >= 2, Mid$(Base32Chars, b + 1, 1), "=")
        sb = sb & IIf(numCharsToOutput >= 3, Mid$(Base32Chars, c + 1, 1), "=")
        sb = sb & IIf(numCharsToOutput >= 4, Mid$(Base32Chars, d + 1, 1), "=")
        sb = sb & IIf(numCharsToOutput >= 5, Mid$(Base32Chars, e + 1, 1), "=")
        sb = sb & IIf(numCharsToOutput >= 6, Mid$(Base32Chars, f + 1, 1), "=")
        sb = sb & IIf(numCharsToOutput >= 7, Mid$(Base32Chars, g + 1, 1), "=")
        sb = sb & IIf(numCharsToOutput >= 8, Mid$(Base32Chars, h + 1, 1), "=")
    Loop
    ToBase32 = sb
End Function

Public Function FromBase32(inputS As String) As Byte()
    Dim output() As Byte, bitIndex As Integer, inputIndex As Integer, outputBits As Integer, outputIndex As Integer, byteIndex As Integer, bits As Integer, i As Long
    inputS = Replace(Trim(UCase$(inputS)), "=", ""): ReDim output(Int(Len(inputS) * 5 / 8) - 1)
    bitIndex = 0: inputIndex = 1: outputBits = 0: outputIndex = LBound(output)
    Do While outputIndex <= UBound(output)
        byteIndex = InStr(1, Base32Chars, Mid$(inputS, inputIndex, 1)) - 1
        bits = IIf(5 - bitIndex < 8 - outputBits, 5 - bitIndex, 8 - outputBits)
		output(outputIndex) = (output(outputIndex) * (2 ^ bits) + (byteIndex \ (2 ^ (5 - (bitIndex + bits))))) And &HFF
        bitIndex = bitIndex + bits
        If bitIndex >= 5 Then bitIndex = 0: inputIndex = inputIndex + 1
        outputBits = outputBits + bits: If outputBits >= 8 Then outputIndex = outputIndex + 1: outputBits = 0
    Loop
    FromBase32 = output
End Function

Private Function GetNextGroup(ByRef inputByte() As Byte, ByRef offset As Long, ByRef a As Byte, ByRef b As Byte, ByRef c As Byte, ByRef d As Byte, ByRef e As Byte, ByRef f As Byte, ByRef g As Byte, ByRef h As Byte) As Integer
    Dim b1 As Long, b2 As Long, b3 As Long, b4 As Long, b5 As Long, retVal As Integer
    Select Case UBound(inputByte) - offset + 1
        Case 1: retVal = 2
        Case 2: retVal = 4
        Case 3: retVal = 5
        Case 4: retVal = 7
        Case Else: retVal = 8
    End Select
    If offset <= UBound(inputByte) Then b1 = inputByte(offset) Else b1 = 0
	offset = offset + 1
    If offset <= UBound(inputByte) Then b2 = inputByte(offset) Else b2 = 0
	offset = offset + 1
    If offset <= UBound(inputByte) Then b3 = inputByte(offset) Else b3 = 0
	offset = offset + 1
    If offset <= UBound(inputByte) Then b4 = inputByte(offset) Else b4 = 0
	offset = offset + 1
    If offset <= UBound(inputByte) Then b5 = inputByte(offset) Else b5 = 0
	offset = offset + 1
    a = (b1 \ 8) And &HFF:    b = (((b1 And 7) * 4) Or (b2 \ 64)) And &HFF
    c = (b2 \ 2) And &H1F:    d = (((b2 And 1) * 16) Or (b3 \ 16)) And &HFF
    e = (((b3 And 15) * 2) Or (b4 \ 128)) And &HFF:    f = (b4 \ 4) And &H1F
    g = (((b4 And 3) * 8) Or (b5 \ 32)) And &HFF:    h = (b5 And 31) And &HFF
    GetNextGroup = retVal
End Function