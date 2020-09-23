Attribute VB_Name = "AND1"

Public Sub D1()
On Error GoTo 4
Randomize Timer
Randomize Rnd
4 I = Abs(RNDD(7, False))    'some possible math functions for use
If I = 1 Then GoTo 3
N1 = Abs(RNDD(N2(I), True))
3 RR0 = Abs(RNDD(RR1(I), True))
If RR0 < 1000000000 Then GoTo 3

For X = 0 To Form3.ScaleWidth
For Y = 0 To Form3.ScaleHeight
XX = Form3.ScaleWidth - X
YY = Form3.ScaleHeight - Y
A1 = ((X And Y) And (XX + YY))
A2 = ((XX And Y) And (X + YY))
A3 = ((X And YY) And (XX + Y))
A4 = ((XX And YY) And (X + Y))


AA1 = (A1 + A2 + A3 + A4)
'AA2 = (A1 Xor A2 Xor A3 Xor A4)
'AA3 = (A1 Or A2 Or A3 Or A4)
'AA4 = (A1 Eqv A2 Eqv A3 Eqv A4)
A = AA1
'N1 = N2(I): RR0 = RR1(I)
If I = 1 Then A = SZ(CIS((A * 1.063) ^ 1.24)) * RR0
If I = 2 Then A = SZ(CIS(((A * 1.026) ^ 1.14) * N1)) * RR0
If I = 3 Then A = SZ(CIS(((A ^ 1.000001) * N1) * N1)) * RR0
If I = 4 Then A = SZ(CIS((A * 1.00000000000011) ^ 6.1 - N1)) * RR0
If I = 5 Then A = SZ(CIS((A * 1.24) ^ 6.01 + N1)) * RR0
If I = 6 Then A = SZ(CIS((A - N1) * N1)) * RR0
If I = 7 Then A = SZ(CIS((((A * 4.2) ^ 1.01) + N1) * N1)) * RR0
A = (A ^ 2)
SetPixel Form3.hdc, XX, Y, A
SetPixel Form3.hdc, X, YY, A          ' .....
SetPixel Form3.hdc, X, Y, A          ' .....
SetPixel Form3.hdc, XX, YY, A          '.....

DoEvents
Next Y, X
2 End Sub



Public Sub D2()
On Error GoTo 4
Randomize Timer
Randomize Rnd
4 'I1 = Abs(RNDD(7, False))   'some possible math functions for use
'I = I + 7

'If I = 1 Then GoTo 3
'N1 = Abs(RNDD(N2(I), True))
'3 RR0 = Abs(RNDD(RR1(I), True))
'If I = 2 And RR0 < 1000000000 Then GoTo 3
3 RR0 = RR1(8) 'Abs(RNDD(RR1(8), True))
If RR0 < 1000000000 Then GoTo 3
For X = 0 To Form3.ScaleWidth
For Y = 0 To Form3.ScaleHeight
XX = Form3.ScaleWidth - X
YY = Form3.ScaleHeight - Y
A1 = ((X And Y) And -(XX + YY))
A2 = ((XX And Y) And -(X + YY))
A3 = ((X And YY) And -(XX + Y))
A4 = ((XX And YY) And -(X + Y))

A5 = ((X And Y) Or -(XX + YY))
A6 = ((XX And Y) Or -(X + YY))
A7 = ((X And YY) Or -(XX + Y))
A8 = ((XX And YY) Or -(X + Y))

AA1 = (A1 + A2 + A3 + A4)
AA2 = (A5 + A6 + A7 + A8)
'AA2 = (A1 Xor A2 Xor A3 Xor A4)
'AA3 = (A1 Or A2 Or A3 Or A4)
'AA4 = (A1 Eqv A2 Eqv A3 Eqv A4)
A = Abs(1.31 * AA1 * AA2) ^ 2.88

'If I1 = 1 Then
A = SZ(CIS(A)) * 63480000000#

'If I1 = 2 Then A = SZ(CIS(A * N1)) * RR0
'If I1 = 3 Then A = SZ(CIS((A * N1) * N1)) * RR0
'If I1 = 4 Then A = SZ(CIS(A - N1)) * RR0
'If I1 = 5 Then A = SZ(CIS(A + N1)) * RR0
'If I1 = 6 Then A = SZ(CIS((A - N1) * N1)) * RR0
'If I1 = 7 Then A = SZ(CIS((A + N1) * N1)) * RR0


A = (A ^ 2)
SetPixel Form3.hdc, XX, Y, A
SetPixel Form3.hdc, X, YY, A ' .....
SetPixel Form3.hdc, X, Y, A ' .....
SetPixel Form3.hdc, XX, YY, A '.....

DoEvents
Next Y, X
2 End Sub

