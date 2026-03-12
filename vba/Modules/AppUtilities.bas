Attribute VB_Name = "AppUtilities"
Option Explicit
Option Private Module

'Public Property Get App() As App
'    Static Result As App
'    If Result Is Nothing Then Set Result = New App
'    Set App = Result.Self
'End Property

Public Function IsInArray(value As String, arr As Variant) As Boolean
    Dim i As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) = value Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Public Function GetFolderPathOf(ByVal FilePath As String) As String
    GetFolderPathOf = Left(FilePath, InStrRev(FilePath, "\") - 1)
End Function

Public Function GetFontOf(ByRef Shape As PowerPoint.Shape) As PowerPoint.Font
    Set GetFontOf = Shape.TextFrame.TextRange.Font
End Function

Private Sub DisplayCustomColorShapesFillRGB()
    Dim TargetSlide As Slide: Set TargetSlide = ActivePresentation.Slides(1)

    Dim RowLetter As Variant
    Dim ColumnNumber As Integer
    Dim ShapeName As String
    Dim ColorRGB As Long
    Dim RedComponent As Long
    Dim GreenComponent As Long
    Dim BlueComponent As Long

    Dim CustomShape As Shape

    For Each RowLetter In Split("A B C D E F G H I J")
        For ColumnNumber = 1 To 5
            ShapeName = "CustomColor" & RowLetter & ColumnNumber

            On Error Resume Next
            Set CustomShape = TargetSlide.Shapes(ShapeName)
            On Error GoTo 0

            If Not CustomShape Is Nothing Then
                If CustomShape.Fill.Visible Then
                    ColorRGB = CustomShape.Fill.ForeColor.Rgb
                    RedComponent = ColorRGB Mod 256
                    GreenComponent = (ColorRGB \ 256) Mod 256
                    BlueComponent = (ColorRGB \ 65536) Mod 256

                    Debug.Print ShapeName & ": RGB(" & RedComponent & ", " & GreenComponent & ", " & BlueComponent & ")"
                Else
                    Debug.Print ShapeName & ": Fill not visible"
                End If
            Else
                Debug.Print ShapeName & ": Not found"
            End If

            Set CustomShape = Nothing
        Next ColumnNumber
    Next RowLetter
End Sub
Sub ResaltarCustomColorShapes()
    Dim sld As Slide
    Set sld = ActiveWindow.View.Slide ' O usa SlideShowWindow si estás en presentación
    
    Dim shp As Shape
    Dim count As Long
    count = 0
    
    For Each shp In sld.Shapes
        If shp.Name Like "CustomColor*" Then
            ' Cambiar a color amarillo (sin usar ppFillNone)
            If shp.Fill.Type <> 0 Then ' 0 = ppFillNone
                shp.Fill.ForeColor.Rgb = Rgb(255, 255, 0)
            End If
            
            shp.ZOrder msoBringToFront
            shp.Select Replace:=False
            
            count = count + 1
        End If
    Next shp
    
    MsgBox count & " forma(s) seleccionada(s).", vbInformation
End Sub

