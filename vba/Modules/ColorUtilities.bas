Attribute VB_Name = "ColorUtilities"
Option Explicit

Public Type Rgb
    R As Integer
    G As Integer
    B As Integer
End Type

Public Type Hsl
    h As Double
    s As Double
    L As Double
End Type

Public Enum ColorModelEnum
    LongModel
    RgbModel
    HexModel
    HslModel
End Enum

Public Sub RunTests()
    TestColor LongColor(vbRed), rgbRed, "FF0000", 255, 0, 0, 0, 1, 0.5
    TestColor LongColor(vbRed), vbRed, "FF0000", 255, 0, 0, 0, 1, 0.5
    TestColor LongColor(vbGreen), vbGreen, "00FF00", 0, 255, 0, 120, 1, 0.5
    TestColor LongColor(vbBlue), vbBlue, "0000FF", 0, 0, 255, 240, 1, 0.5
    TestColor LongColor(vbYellow), vbYellow, "FFFF00", 255, 255, 0, 60, 1, 0.5
    TestColor LongColor(vbCyan), vbCyan, "00FFFF", 0, 255, 255, 180, 1, 0.5
    TestColor LongColor(vbMagenta), vbMagenta, "FF00FF", 255, 0, 255, 300, 1, 0.5
    TestColor LongColor(vbBlack), vbBlack, "000000", 0, 0, 0, 0, 0, 0
    TestColor LongColor(vbWhite), vbWhite, "FFFFFF", 255, 255, 255, 0, 0, 1
    
    TestColor RgbColor(255, 99, 71), &HFF6347, "FF6347", 255, 99, 71, 9, 1, 0.64
    TestColor RgbColor(70, 130, 180), &H4682B4, "4682B4", 70, 130, 180, 207, 0.44, 0.49
    TestColor RgbColor(218, 112, 214), &HDA70D6, "DA70D6", 218, 112, 214, 302, 0.59, 0.65
    TestColor RgbColor(50, 205, 50), &H32CD32, "32CD32", 50, 205, 50, 120, 0.61, 0.5
    TestColor RgbColor(255, 215, 0), &HFFD700, "FFD700", 255, 215, 0, 51, 1, 0.5
    TestColor RgbColor(255, 69, 0), &HFF4500, "FF4500", 255, 69, 0, 16, 1, 0.5
    TestColor RgbColor(46, 139, 87), &H2E8B57, "2E8B57", 46, 139, 87, 146, 0.5, 0.36
    TestColor RgbColor(106, 90, 205), &H6A5ACD, "6A5ACD", 106, 90, 205, 248, 0.53, 0.58
    TestColor RgbColor(255, 105, 180), &HFF69B4, "FF69B4", 255, 105, 180, 330, 1, 0.71
    TestColor RgbColor(138, 43, 226), &H8A2BE2, "8A2BE2", 138, 43, 226, 271, 0.76, 0.53
    TestColor RgbColor(127, 255, 0), &H7FFF00, "7FFF00", 127, 255, 0, 90, 1, 0.5
    TestColor RgbColor(210, 105, 30), &HD2691E, "D2691E", 210, 105, 30, 25, 0.75, 0.47
    TestColor RgbColor(100, 149, 237), &H6495ED, "6495ED", 100, 149, 237, 219, 0.79, 0.66
    TestColor RgbColor(220, 20, 60), &HDC143C, "DC143C", 220, 20, 60, 348, 0.83, 0.47
    TestColor RgbColor(255, 140, 0), &HFF8C00, "FF8C00", 255, 140, 0, 33, 1, 0.5
    TestColor RgbColor(139, 0, 0), &H8B0000, "8B0000", 139, 0, 0, 0, 1, 0.27
    
    TestColor HexColor("E9967A"), &HE9967A, "E9967A", 233, 150, 122, 15, 0.72, 0.7
    TestColor HexColor("8FBC8F"), &H8FBC8F, "8FBC8F", 143, 188, 143, 120, 0.25, 0.65
    TestColor HexColor("483D8B"), &H483D8B, "483D8B", 72, 61, 139, 248, 0.39, 0.39
    TestColor HexColor("2F4F4F"), &H2F4F4F, "2F4F4F", 47, 79, 79, 180, 0.25, 0.25
    TestColor HexColor("00CED1"), &HCED1, "00CED1", 0, 206, 209, 181, 1, 0.41
    TestColor HexColor("9400D3"), &H9400D3, "9400D3", 148, 0, 211, 282, 1, 0.41
    TestColor HexColor("FF1493"), &HFF1493, "FF1493", 255, 20, 147, 327, 1, 0.54
    TestColor HexColor("00BFFF"), &HBFFF, "00BFFF", 0, 191, 255, 195, 1, 0.5
    TestColor HexColor("696969"), &H696969, "696969", 105, 105, 105, 0, 0, 0.41
    TestColor HexColor("1E90FF"), &H1E90FF, "1E90FF", 30, 144, 255, 210, 1, 0.56
    TestColor HexColor("B22222"), &HB22222, "B22222", 178, 34, 34, 0, 0.68, 0.42
    TestColor HexColor("FFFAF0"), &HFFFAF0, "FFFAF0", 255, 250, 240, 40, 1, 0.97
    TestColor HexColor("228B22"), &H228B22, "228B22", 34, 139, 34, 120, 0.61, 0.34
    TestColor HexColor("FF00FF"), &HFF00FF, "FF00FF", 255, 0, 255, 300, 1, 0.5
    TestColor HexColor("DCDCDC"), &HDCDCDC, "DCDCDC", 220, 220, 220, 0, 0, 0.86
    
    TestColor HslColor(240, 1, 0.5), &HFF, "0000FF", 0, 0, 255, 240, 1, 0.5
    TestColor HslColor(0, 1, 0.5), (vbRed), "FF0000", 255, 0, 0, 0, 1, 0.5
    TestColor HslColor(120, 1, 0.5), (vbGreen), "00FF00", 0, 255, 0, 120, 1, 0.5
    TestColor HslColor(240, 1, 0.5), (vbBlue), "0000FF", 0, 0, 255, 240, 1, 0.5
    TestColor HslColor(60, 1, 0.5), (vbYellow), "FFFF00", 255, 255, 0, 60, 1, 0.5
    TestColor HslColor(180, 1, 0.5), (vbCyan), "00FFFF", 0, 255, 255, 180, 1, 0.5
    TestColor HslColor(300, 1, 0.5), (vbMagenta), "FF00FF", 255, 0, 255, 300, 1, 0.5
    TestColor HslColor(0, 0, 0), (vbBlack), "000000", 0, 0, 0, 0, 0, 0
    TestColor HslColor(0, 0, 1), (vbWhite), "FFFFFF", 255, 255, 255, 0, 0, 1
End Sub
Private Sub TestColor(Sut As Color, LongColor As Long, _
                                    HexColor As String, _
                                    RgbColorR As Integer, RgbColorG As Integer, RgbColorB As Integer, _
                                    HslColorH As Double, HslColorS As Double, HslColorL As Double)
'    Debug.Assert Sut.AsLong = LongColor
'    Debug.Assert Sut.AsHex = HexColor
    Debug.Assert AreRgbEquivalent(Sut.AsRgb, RgbColorR, RgbColorG, RgbColorB)
    Debug.Assert AreHslEquivalent(Sut.AsHsl, HslColorH, HslColorS, HslColorL)
End Sub
Private Function AreRgbEquivalent(value As Rgb, R As Integer, G As Integer, ByVal B As Integer) As Boolean
    AreRgbEquivalent = (value.R = R) And (value.G = G) And (value.B = B)
End Function
Private Function AreHslEquivalent(value As Hsl, h As Double, s As Double, ByVal L As Double) As Boolean
    AreHslEquivalent = (value.h - h < 1) And (value.s - s < 0.01) And (value.L - L < 0.01)
End Function


Public Function RgbColor(ByVal R As Integer, ByVal G As Integer, ByVal B As Integer) As Color
    Dim Result As New Color
    Result.Initialize RgbModel, R, G, B
    
    Set RgbColor = Result
End Function

Public Function LongColor(ByVal value As Long) As Color
    Dim Result As New Color
    Result.Initialize LongModel, value
    
    Set LongColor = Result
End Function

Public Function HslColor(ByVal h As Double, ByVal s As Double, ByVal L As Double) As Color
    Dim Result As New Color
    Result.Initialize HslModel, h, s, L
    
    Set HslColor = Result
End Function

Public Function HexColor(ByVal value As String) As Color
    Dim Result As New Color
    Result.Initialize HexModel, value
    
    Set HexColor = Result
End Function

