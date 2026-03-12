VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private IsInitialized As Boolean
Private pLongColor As Long

Public Sub Initialize(ByVal Model As ColorModelEnum, ParamArray value() As Variant)
    If IsInitialized Then Err.Raise 17
    
    Select Case True
        Case Model = LongModel:             InitializeWithLong value(0)
        Case Model = HexModel:              InitializeWithHex value(0)
        Case Model = RgbModel:              InitializeWithRgb value(0), value(1), value(2)
        Case Model = HslModel:              InitializeWithHsl value(0), value(1), value(2)
    End Select
    
    IsInitialized = True
End Sub
Private Sub InitializeWithLong(ByVal Color As Long)
    pLongColor = Color
End Sub
Private Sub InitializeWithRgb(ByVal R As Integer, ByVal G As Integer, ByVal B As Integer)
    pLongColor = VBA.Rgb(R, G, B)
End Sub
Private Sub InitializeWithHex(ByVal value As String)
    If Not value Like "??????" Then Err.Raise 17

    Dim Color As Rgb
    Color.R = Val("&H" & (Mid(value, 1, 2)))
    Color.G = Val("&H" & (Mid(value, 3, 2)))
    Color.B = Val("&H" & (Mid(value, 5, 2)))

    pLongColor = RgbColor(Color.R, Color.G, Color.B).AsLong
End Sub
Private Sub InitializeWithHsl(ByVal h As Double, ByVal s As Double, ByVal L As Double)
    Dim Color As Rgb
    If s = 0 Then 'Gray
        Color.R = L * 255
        Color.G = L * 255
        Color.B = L * 255
    Else
        Dim q As Double
        
        Select Case True
            Case L < 1 / 2:         q = (L * (1 + s))
            Case Else:              q = (L * (1 - s)) + s
        End Select

        Color.R = HueToRGB(2 * L - q, q, h / 360 + 1 / 3) * 255
        Color.G = HueToRGB(2 * L - q, q, h / 360 + 0 / 3) * 255
        Color.B = HueToRGB(2 * L - q, q, h / 360 - 1 / 3) * 255
    End If

    pLongColor = RgbColor(Color.R, Color.G, Color.B).AsLong
End Sub
Private Function HueToRGB(p As Double, q As Double, t As Double) As Double
    If t < 0 Then t = t + 1
    If t > 1 Then t = t - 1
    If t < 1 / 6 Then HueToRGB = p + (q - p) * 6 * t
    If t >= 1 / 6 And t < 1 / 2 Then HueToRGB = q
    If t >= 1 / 2 And t < 2 / 3 Then HueToRGB = p + (q - p) * (2 / 3 - t) * 6
    If t >= 2 / 3 Then HueToRGB = p
End Function

Public Property Get AsLong() As Long
    AsLong = pLongColor
End Property

Public Property Get AsRgb() As ColorUtilities.Rgb
    Dim Result As ColorUtilities.Rgb
    
    Result.R = (AsLong \ (2 ^ 0)) Mod (2 ^ 8)
    Result.G = (AsLong \ (2 ^ 8)) Mod (2 ^ 8)
    Result.B = (AsLong \ (2 ^ 16)) Mod (2 ^ 8)
    
    AsRgb = Result
End Property

Public Property Get AsHsl() As ColorUtilities.Hsl
    Dim Norms(1 To 3) As Double
    With AsRgb
        Norms(1) = .R / 255
        Norms(2) = .G / 255
        Norms(3) = .B / 255
    End With
    
    Dim Delta As Double
    Delta = Max(Norms) - Min(Norms)

    With AsHsl
                                                        .L = (Max(Norms) + Min(Norms)) / 2
        
        Select Case True
            Case Delta = 0:                             .s = 0
            Case Else
                Select Case True
                    Case .L < 1 / 2:                    .s = Delta / (Max(Norms) + Min(Norms))
                    Case Else:                          .s = Delta / (2 - Max(Norms) - Min(Norms))
                End Select
        End Select
        
        Select Case True
            Case Delta = 0:                             .h = 0
            Case Else:
                Select Case True
                    Case Max(Norms) = Norms(1)
                        Select Case True
                            Case Norms(2) < Norms(3):   .h = (Norms(2) - Norms(3)) / Delta + 6
                            Case Else:                  .h = (Norms(2) - Norms(3)) / Delta + 0
                        End Select
                    Case Max(Norms) = Norms(2):         .h = (Norms(3) - Norms(1)) / Delta + 2
                    Case Max(Norms) = Norms(3):         .h = (Norms(1) - Norms(2)) / Delta + 4
                End Select
                                                        .h = 360 * (.h / 6)
        End Select
    End With
End Property

Public Property Get AsHex() As String
    AsHex = Right("0" & Hex(AsRgb.R), 2) & Right("0" & Hex(AsRgb.G), 2) & Right("0" & Hex(AsRgb.B), 2)
End Property

Public Function ToGray(Optional ByVal Rf As Double = 0.299, Optional ByVal Gf As Double = 0.587, Optional ByVal Bf As Double = 0.114) As Color
    Dim value As Integer
    value = Rf * AsRgb.R + Gf * AsRgb.G + Bf * AsRgb.B

    Set ToGray = RgbColor(value, value, value)
End Function

Public Function ToDark(Optional ByVal Factor As Double = 0.8) As Color
    If Factor < 0 Or Factor > 1 Then Err.Raise 17

    Set ToDark = RgbColor(Factor * AsRgb.R, Factor * AsRgb.G, Factor * AsRgb.B)
End Function

Public Function ToLightness(ByVal value As Double) As Color
    Dim NewValue As Double
    NewValue = value
    
    Select Case True
        Case NewValue > 1: NewValue = 1
        Case NewValue < 0: NewValue = 0
    End Select
    
    Set ToLightness = HslColor(AsHsl.h, AsHsl.s, NewValue)
End Function

Public Function ChangeLightnessBy(ByVal value As Double) As Color
    Set ChangeLightnessBy = ToLightness(AsHsl.L + value)
End Function

Public Function InvertLightness() As Color
    Set InvertLightness = ToLightness(1 - AsHsl.L)
End Function

Public Function ToOposite() As Color
    Set ToOposite = RgbColor(255 - AsRgb.R, 255 - AsRgb.G, 255 - AsRgb.B)
End Function

Public Sub Log()
    Debug.Print "AsLong:" & vbTab & AsLong
    Debug.Print "AsHex:" & vbTab & AsHex
    Debug.Print "AsRgb:" & vbTab & AsRgb.R; AsRgb.G; AsRgb.B
    Debug.Print "AsHsl:" & vbTab & Round(AsHsl.h); Round(AsHsl.s, 2); Round(AsHsl.L, 2)
End Sub
