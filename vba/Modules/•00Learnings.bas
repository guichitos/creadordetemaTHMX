Attribute VB_Name = "•00Learnings"
Option Explicit
Option Private Module




Private Sub LearnToSaveEntireApp()
    Dim Presentation As PowerPoint.Presentation
    Set Presentation = Slide56.Parent
    
    Presentation.Save
    
'    FileUtilities.CopyFolder GetFolderPathOf(FileUtilities.GetLocalPath(Presentation.FullName)), GetKnownFolderPath(kfID_Downloads) & "\Creator - " & Format(Now, "YYYY.MM.DD.hhnn")
    FileUtilities.CopyFolder ParentFolder(GetLocalPath(Presentation.FullName)), GetKnownFolderPath(kfID_Downloads) & "\Creator - " & Format(Now, "YYYY.MM.DD.hhnn")
End Sub

Private Sub LearnWhichAlTTextIs()
    Dim CommandBar As CommandBar
    
    
End Sub

Private Sub LearnToSelectEntireAlternativePaletteShapes()
    Dim Range As PowerPoint.ShapeRange
    
    Dim i As Integer
    i = 3
    
    Set Range = App.Slide.Shapes.Range(Array("CopyCurrentColorPalette" & i, "PaletteColorA" & i))
    
    Range.Visible = False
    
    Range.Visible = True
End Sub


Private Sub LearnToHaveShapesInGroup()
    Dim Slide As PowerPoint.Slide
    Set Slide = Slide56
    
    Dim Shape As Shape
    Set Shape = Slide.Shapes("PaletteColor1Group")
    
End Sub


Private Sub LearnAboutCollections()
    Dim Col As New Collection
    
    Col.Add 7
    Col.Add 8
    Col.Add 9
    
    Debug.Print Col(1), Col(Col.count)
End Sub

Private Sub LearnToGrabColorsFromShape()
    Dim Shape As PowerPoint.Shape
    Set Shape = ActiveWindow.Selection.ShapeRange.Item(1)
    
    With Shape
        Dim Color As Variant
        Color = .Fill.ForeColor
        With .TextFrame
            .TextRange.Text = GetHexFromLong((Color))
            .AutoSize = ppAutoSizeNone
            .MarginBottom = 0
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .WordWrap = msoTrue
            .HorizontalAnchor = msoAnchorCenter
        End With
    End With
End Sub
Function GetHexFromLong(Color As Long) As String 'https://exceloffthegrid.com/convert-color-codes/
    Dim R As String
    Dim G As String
    Dim B As String
    
    R = Format(Hex((Color \ (2 ^ 0)) Mod (2 ^ 8)), "00")
    G = Format(Hex((Color \ (2 ^ 8)) Mod (2 ^ 8)), "00")
    B = Format(Hex((Color \ (2 ^ 16)) Mod (2 ^ 8)), "00")
    
    GetHexFromLong = "#" & R & G & B
End Function
Function GetHexFromRGB(Red As Integer, Green As Integer, Blue As Integer) As String
    GetHexFromRGB = "#" & VBA.Right$("00" & VBA.Hex(Red), 2) & _
                            VBA.Right$("00" & VBA.Hex(Green), 2) & _
                            VBA.Right$("00" & VBA.Hex(Blue), 2)

End Function
Function GetRgbFromLong(LongColor As Long, Rgb As String) As Integer
    Select Case Rgb
        Case "R"
            GetRgbFromLong = (LongColor \ (2 ^ 0)) Mod (2 ^ 8)
        Case "G"
            GetRgbFromLong = (LongColor \ (2 ^ 8)) Mod (2 ^ 8)
        Case "B"
            GetRgbFromLong = (LongColor \ (2 ^ 16)) Mod (2 ^ 8)
    End Select
End Function

Public Sub Eyedropper()
    Dim CommandBar As CommandBar
    For Each CommandBar In Application.CommandBars
        On Error Resume Next
        CommandBar.ShowPopup
        On Error GoTo 0
        
        Dim control As CommandBarControl
        For Each control In CommandBar.Controls
            Debug.Print control.Id; control.Caption
        Next
    Next
End Sub

Private Sub FindSetDefaultShapeCommand()
    Dim i As Long
    For i = 1 To 5000
        Dim control As CommandBarControl
        Set control = PowerPoint.Application.CommandBars.FindControl(Id:=i)
        
        If Not control Is Nothing Then
        Debug.Assert Not control.Caption Like "*icture*"
        Debug.Print control.Caption, i
        End If
    Next
End Sub

Private Sub LearnToRetrieveThemeEffects()
    With App.Presentation.Slides(1).Master.Theme
        Debug.Print .ThemeEffectScheme.Creator
    End With
End Sub

Private Sub LearnAboutThemeColorSchemeXML()
    With App.Presentation.Slides(1).Master.Theme
        .ThemeColorScheme.Save Environ("USERPROFILE") & "\Downloads\ThemeColors.xml"
    End With
End Sub

Private Sub LearnAboutXmlParts()
    With App.Presentation.CustomXMLParts
        
    End With
End Sub
