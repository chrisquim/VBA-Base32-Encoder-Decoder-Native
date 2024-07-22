'VBA BASE32 ENCODER DECODER NATIVE
'Develop converted base from .Net Code C# Source Code
'https://github.com/dotnet/aspnetcore/blob/01cc669960821e23ef3275cd5ad81f7192972010/src/Identity/Extensions.Core/src/Base32.cs
'Minified Intended for direct use.

Option Explicit

'https://datatracker.ietf.org/doc/html/rfc3548#section-5
'https://en.wikipedia.org/wiki/Base32
Private Const Base32Chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"

Sub SampleTest() 'FOR TESTING. NOT REQUIRED
    Dim arrB() As Byte, s As String, n As Long, i As Long, arrD() As Byte
    n = 21      'ARBITRARY BYTE ARRAY LENGTH
    ReDim arrB(n)
    For i = 0 To n: arrB(i) = CByte(Rnd() * 255): Next
    s = Base32Encode(arrB)
    arrD = Base32Decode(s)
    For i = 0 To n: Debug.Print arrB(i);: Next: Debug.Print vbNullString
    For i = 0 To n: Debug.Print arrD(i);: Next: Debug.Print vbNullString
    Debug.Print s
End Sub

Public Function Base32Encode$(iB() As Byte)
    'Please dont forget the 1 Private Helper Function GNG%
    Dim sb$, o&, n&, i&, x(7) As Byte: o = LBound(iB): Do While o <= UBound(iB): n = GNG(iB, o, x): For i = 0 To 7: sb = sb & IIf(n >= i + 1, Mid$(Base32Chars, x(i) + 1, 1), "="): Next: Loop: Base32Encode = sb
End Function

Public Function Base32Decode(ByVal str As String) As Byte()
    'No dependency necessary
    Dim bI&, iI&, oB&, oI&, byI&, b&, i&, o() As Byte: str = Replace$(Trim$(UCase$(str)), "=", ""): ReDim o(Int(Len(str) * 5 / 8) - 1): bI = 0: iI = 1: oB = 0: oI = LBound(o)
    Do While oI <= UBound(o): byI = InStr(1, Base32Chars, Mid$(str, iI, 1)) - 1: b = IIf(5 - bI < 8 - oB, 5 - bI, 8 - oB): o(oI) = (o(oI) * (2 ^ b) + (byI \ (2 ^ (5 - (bI + b))))) And &HFF: bI = bI + b: If bI >= 5 Then bI = 0: iI = iI + 1
        oB = oB + b: If oB >= 8 Then oI = oI + 1: oB = 0
    Loop: Base32Decode = o
End Function

'HELPER PRIVATE FUNCTION for Base32Encode$
Private Function GNG%(ByRef iB() As Byte, ByRef of&, xI() As Byte)
    Dim i&, r&, x&(4): Select Case UBound(iB) - of + 1: Case 1: r = 2: Case 2: r = 4: Case 3: r = 5: Case 4: r = 7: Case Else: r = 8: End Select
    For i = 0 To 4: If of <= UBound(iB) Then x(i) = iB(of) Else x(i) = 0
        of = of + 1: Next: xI(0) = (x(0) \ 8) And &HFF: xI(1) = (((x(0) And 7) * 4) Or (x(1) \ 64)) And &HFF: xI(2) = (x(1) \ 2) And &H1F: xI(3) = (((x(1) And 1) * 16) Or (x(2) \ 16)) And &HFF
    xI(4) = (((x(2) And 15) * 2) Or (x(3) \ 128)) And &HFF: xI(5) = (x(3) \ 4) And &H1F: xI(6) = (((x(3) And 3) * 8) Or (x(4) \ 32)) And &HFF: xI(7) = (x(4) And 31) And &HFF: GNG = r
End Function
