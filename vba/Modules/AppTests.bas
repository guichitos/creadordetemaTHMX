Attribute VB_Name = "AppTests"
Option Explicit
Option Private Module

Private Sub TestToSetAnAlternativeColorPalette()
    With New App
        Dim PlainPalette As OfficeUtilities.OfficeColorPalette
        .SetAlternativePalette 2, PlainPalette
        
        .SetAlternativePalette 3, GetColorPalette("Rgb palette", rgbAqua, rgbAquamarine, rgbBlack, rgbBlueViolet, rgbCadetBlue, rgbCornflowerBlue, rgbCornsilk, rgbCoral, rgbDarkGrey, rgbDarkCyan, rgbDeepSkyBlue, rgbBlue, DefaultColorMapping)
    End With
End Sub

Private Sub TestColorClass()
    Dim Color As New Color

'    Color.Initialize (rgbBlue)
'
'    Debug.Assert Color.AsRgb.R = 0
'    Debug.Assert Color.AsRgb.G = 0
'    Debug.Assert Color.AsRgb.B = 255
'
'    Debug.Assert Color.AsHex = "0000FF"
'
'    Dim AnotherColor As New Color
'    AnotherColor.Initialize 255, 0, 0
'
'    Debug.Assert AnotherColor.AsRgb.R = 255
'    Debug.Assert AnotherColor.AsRgb.G = 0
'    Debug.Assert AnotherColor.AsRgb.B = 0
'
'    Debug.Assert AnotherColor.AsHex = "FF0000"
'
'    Dim YetAnotherColor As New Color
'    YetAnotherColor.Initialize "FFFF00"
'
'    Debug.Assert YetAnotherColor.AsRgb.R = 255
'    Debug.Assert YetAnotherColor.AsRgb.G = 255
'    Debug.Assert YetAnotherColor.AsRgb.B = 0
'
'    Debug.Assert YetAnotherColor.AsLong = Rgb(255, 255, 0)
    Set Color = LongColor(rgbBlue)
    
    Debug.Print Color.AsHsl.h, Color.AsHsl.s,
    
End Sub

