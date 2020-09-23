Attribute VB_Name = "Module3"
Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Declare Function hiword Lib "TLBINF32" (ByVal DWord As Long) As Integer
Public Declare Function loword Lib "TLBINF32" (ByVal DWord As Long) As Integer
Public Declare Function hiByte Lib "TLBINF32" Alias "hibYte" (ByVal Word As Integer) As Byte
Public Declare Function loByte Lib "TLBINF32" Alias "lobYte" (ByVal Word As Integer) As Byte


Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long


Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long


Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Const pi = 3.14159265358979

Public Pau As Boolean, A As Single, N2(70) As Double, RR1(70) As Double
Public I, X As Long, Y As Long, XX As Long, YY As Long
Public AB(90) As Single
Public N1 As Single, R1 As RECT


Public Function CIS(Num) As Double ' Math Combinations for the Pictures
Dim Rev As Double
Rev = Sin(Num) * Cos(Num) * Sin(Cos(Num)) * Cos(Sin(Num)) ' -1=<Rev<=1
CIS = Atn(Rev) * Rev '>=1 or <=-1
End Function

Public Function SZ(Num) As Double  'Math Combinations for the Pictures
Dim Rev As Double
Rev = Sin(Num) * Cos(Num) * Sin(Cos(Num)) * Cos(Sin(Num)) * Tan(Num) '- 1 >= Rev <= 1
SZ = Atn(Rev) * Log(Abs(Num) + 1) * Sqr(Abs(Num)) * Exp(Rev) * Rev '>=1 or <=-1
End Function

Public Function QW(Num) As Double 'Math Combinations for the Pictures
Dim Rev As Double
Rev = Sin(Num) * Cos(Num) * Sin(Cos(Num)) * Cos(Sin(Num)) ' -1=<Rev<=1
QW = Atn(Rev) * Tan(Rev) * Log(Abs(Rev) + 1) * Sqr(Abs(Rev)) * Exp(Rev) * Sqr(Abs(Num)) * Log(Abs(Num) + 1) '>=1 or <=-1
End Function

Function SMN(Number) '
SMN = Sqr(Abs(-Number * Not Number)) ' = squar root(X*(X+1))
End Function

Public Function Red(Color As Long) As Long 'Value of Red color
Red = loByte(loword(Color))
End Function

Public Function Green(Color As Long) As Long 'Value of Green color
Green = hiByte(loword(Color))
End Function

Public Function Blue(Color As Long) As Long ' Value of Blue color
Blue = hiword(Color)
End Function

Public Sub SLEEP() ' Long Pause function
For G = 1 To 10 ^ 8: Next G
End Sub

Public Sub SLEEP1() ' Short Pause function
For G = 1 To 10 ^ 7.9: Next G
End Sub

Public Function Max(ParamArray Num()) ' Maximum function
Max = Num(0)
For Rev = 1 To UBound(Num)
If Num(Rev) > Max Then Max = Num(Rev)
Next
End Function


Public Function Min(ParamArray Num()) ' Minimu, function
Min = Num(0)
For Rev = 1 To UBound(Num)
If Num(Rev) < Min Then Min = Num(Rev)
Next
End Function

Public Function RNDD(Num2 As Double, Dot As Boolean) As Single
'Function makes a random number greater then 1 and less then -1
Randomize Timer
Randomize Rnd
Dim Rndd1 As Integer
If Abs(Num2) = 1 Then
RNDD = 1
End If
2 RNDD = (Rnd * Num2)
If Abs(RNDD) < 1 Or Abs(RNDD) > Num2 Then GoTo 2
If Dot = False Then RNDD = Int(RNDD)
End Function

Public Function Sgn1(ParamArray Num2() As Variant) As Variant
'The function gets sign of Number by Multiplicating the signs of the wanted numbers
Sgn1 = 1
For Rev = LBound(Num2) To UBound(Num2)
If Sgn(Num2(Rev)) <> 0 Then Sgn1 = Sgn1 * Sgn(Num2(Rev))
Next
End Function

Public Function Sgn2(ParamArray Num2() As Variant) As Variant
'The function gets sign of Number by comparing the Maximum numberes of signs
Sgn2 = 1
For Rev = LBound(Num2) To UBound(Num2)
If Sgn(Num2(Rev)) = -1 Then Rev1 = Rev1 + 1
If Sgn(Num2(Rev)) = 1 Then Rev2 = Rev2 + 1
Next
If Rev1 > Rev2 Then Sgn2 = -1
End Function

Public Function I2()
If I Mod 7 = 0 Then I2 = I / 7 - 1
If I Mod 7 <> 0 Then I2 = (I - ((I Mod 7))) / 7
End Function
