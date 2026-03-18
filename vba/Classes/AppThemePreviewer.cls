VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppThemePreviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private PreviousSelection As PowerPoint.Selection

Private Const ORIGINAL_THEME_FILE_NAME As String = "OriginalTheme.thmx"
Private Const GENERATED_THEME_FILE_NAME As String = "Default theme.thmx"


Private Sub •00InstatiationProcedures():                                                                            End Sub

Private Sub •00EventsProcedures():                                                                                  End Sub

Private Sub •00PublicProcedures():                                                                                  End Sub

Private Sub •00FriendProcedures():                                                                                  End Sub

Friend Sub PreviewTheme()
    ActiveWindow.Selection.Unselect
        
    SyncLayoutsFromSampleSlidesShapesAndPlaceholders
    
    If App.IsWaitingForLayoutClick = True Then Exit Sub
    ApplyStoredLanguageToSlideMaster
    ApplySlideSizeFromMaster
    ApplyDesiredFormatToAllDesignsUsingVba
    DeleteWorkingEnvironment
    SaveToTemporaryFolderAsOriginalTheme
    If Not PythonLibraryIsInstalled Then InstallPythonLibrary
    CreateOfficeThemeBasedOnOriginalThemeAndCallItCustomThemeUsingPython
    Dim PreviousDesignName As String: PreviousDesignName = App.Slide.Design.Name
    Dim PreviousSlidemaster As Design: Set PreviousSlidemaster = GetDesignByName(App.Presentation, PreviousDesignName)
    Dim CustomThemePath As String: CustomThemePath = App.WorkingEnvironmentPath & GENERATED_THEME_FILE_NAME
    PreviousSlidemaster.Preserved = msoFalse
    ClearNonPlaceholderShapesFromLayoutsInLayoutDesignSection 'Otherwise shapes in layouts are duplicated
    App.Presentation.ApplyTheme CustomThemePath
    ResetValuesForLayoutFlow
    App.ThemeFontsHeadingsFontShape.TextFrame.TextRange.Text = App.HeadingsAlternativeFontNameTextbox.TextFrame.TextRange.Text
    App.ThemeFontsBodyFontShape.TextFrame.TextRange.Text = App.BodyAlternativeFontNameTextbox.TextFrame.TextRange.Text
    'RemoveBrandLogoImagesFromOriginalTheme 'Must be deleted
End Sub
Public Sub RemoveBrandLogoImagesFromOriginalTheme()

    On Error GoTo CleanExit
    If App Is Nothing Then Set App = New App
    If App.Presentation Is Nothing Then Exit Sub

    Dim TargetDesign As Design: Set TargetDesign = GetDesignByName(App.Presentation, "OriginalTheme")
    If TargetDesign Is Nothing Then Exit Sub

    Dim TargetMaster As Master: Set TargetMaster = TargetDesign.SlideMaster
    If TargetMaster Is Nothing Then Exit Sub

    RemoveBrandLogoImagesFromShapes TargetMaster.Shapes

    Dim Layout As CustomLayout
    For Each Layout In TargetMaster.CustomLayouts
        RemoveBrandLogoImagesFromShapes Layout.Shapes
    Next Layout

CleanExit:
End Sub


Private Function GetDesignByName(ByVal TargetPresentation As Presentation, ByVal DesignName As String) As Design

    On Error GoTo CleanExit

    Dim CurrentDesign As Design
    For Each CurrentDesign In TargetPresentation.Designs
        If StrComp(CurrentDesign.Name, DesignName, vbTextCompare) = 0 Then
            Set GetDesignByName = CurrentDesign
            Exit Function
        End If
    Next CurrentDesign

CleanExit:
End Function


Private Sub RemoveBrandLogoImagesFromShapes(ByVal TargetShapes As Shapes)

    On Error GoTo CleanExit

    Dim ShapeIndex As Long
    For ShapeIndex = TargetShapes.count To 1 Step -1
        Dim CurrentShape As Shape: Set CurrentShape = TargetShapes(ShapeIndex)
        RemoveBrandLogoImagesFromShape CurrentShape
    Next ShapeIndex

CleanExit:
End Sub


Private Sub RemoveBrandLogoImagesFromShape(ByVal TargetShape As Shape)

    On Error GoTo CleanExit

    If StrComp(TargetShape.Name, "BrandLogoImage", vbBinaryCompare) = 0 Then
        On Error Resume Next
        TargetShape.Delete
        On Error GoTo 0
        Exit Sub
    End If

    If TargetShape.Type = msoGroup Then

        Dim GroupIndex As Long
        For GroupIndex = TargetShape.GroupItems.count To 1 Step -1
            Dim GroupItem As Shape: Set GroupItem = TargetShape.GroupItems(GroupIndex)
            RemoveBrandLogoImagesFromShape GroupItem
        Next GroupIndex

    End If

CleanExit:
End Sub
Public Sub ResetValuesForLayoutFlow()
    App.CurrentSlideIndex = 0
End Sub

Private Sub •00PrivateProcedures():                                                                                 End Sub

Private Sub DeleteWorkingEnvironment()

    If Len(App.WorkingEnvironmentPath) = 0 Then Exit Sub
    
    If Dir(App.WorkingEnvironmentPath, vbDirectory) = "" Then Exit Sub
    
    If InStr(1, App.WorkingEnvironmentPath, "ToolEnvironment", vbTextCompare) = 0 Then
        MsgBox "Safety stop: Not a valid tool environment folder.", vbCritical
        Exit Sub
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error GoTo CleanFail
    fso.DeleteFolder App.WorkingEnvironmentPath, True 'True = force delete
    On Error GoTo 0
    
    App.WorkingEnvironmentPath = ""
    
    Exit Sub

CleanFail:
    'MsgBox "Error deleting working environment.", vbCritical

End Sub
Private Sub ApplyDesiredFormatToAllDesignsUsingVba()
    Dim Design As PowerPoint.Design
    
    For Each Design In App.Presentation.Designs
        ApplyFontsToTheme Design.SlideMaster
        ApplyColorsToTheme Design.SlideMaster
        
        SetDefultShapesInCurrentTheme Design.SlideMaster
    Next Design
End Sub






Private Sub ApplyFontsToTheme(ByRef Master As PowerPoint.Master)
    With Master.Theme.ThemeFontScheme
        .MajorFont.Item(msoThemeLatin).Name = App.HeadingsAlternativeFontNameTextbox.TextFrame.TextRange.Text
        
        .MinorFont.Item(msoThemeLatin).Name = App.BodyAlternativeFontNameTextbox.TextFrame.TextRange.Text
    End With
End Sub

Private Sub ApplyColorsToTheme(ByRef Master As PowerPoint.Master)

    With Master.Theme.ThemeColorScheme
        'Assuming Heading texts use brand color, which is very common
        'this makes the first column of shapes in gallery more useful

        .Colors(msoThemeLight1) = App.PaletteShapes(0)(1).Fill.ForeColor
        .Colors(msoThemeDark1) = App.PaletteShapes(0)(2).Fill.ForeColor
        
        .Colors(msoThemeLight2) = App.PaletteShapes(0)(3).Fill.ForeColor
        .Colors(msoThemeDark2) = App.PaletteShapes(0)(4).Fill.ForeColor
        '/**
        .Colors(msoThemeAccent1) = App.PaletteShapes(0)(5).Fill.ForeColor
        .Colors(msoThemeAccent2) = App.PaletteShapes(0)(6).Fill.ForeColor
        .Colors(msoThemeAccent3) = App.PaletteShapes(0)(7).Fill.ForeColor
        .Colors(msoThemeAccent4) = App.PaletteShapes(0)(8).Fill.ForeColor
        .Colors(msoThemeAccent5) = App.PaletteShapes(0)(9).Fill.ForeColor
        .Colors(msoThemeAccent6) = App.PaletteShapes(0)(10).Fill.ForeColor
        
        .Colors(msoThemeHyperlink) = App.PaletteShapes(0)(11).Fill.ForeColor
        .Colors(msoThemeFollowedHyperlink) = App.PaletteShapes(0)(12).Fill.ForeColor
    End With
End Sub

Private Sub SetDefultShapesInCurrentTheme(Master As PowerPoint.Master)
    PowerPointUtilities.SetMastersDefaultShape App.DefaultThemeShape, Master
    PowerPointUtilities.SetMastersDefaultShape App.DefaultThemeTextBox, Master
    PowerPointUtilities.SetMastersDefaultShape App.DefaultThemeLine, Master
End Sub

Private Sub •00PythonSection():                                                                          End Sub

Private Function PythonLibraryIsInstalled() As Boolean
    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    
    Dim ShellInstrution As String
    ShellInstrution = "python """ & App.localPath & "\Scripts\verify_library.py" & """"
    Dim ShellResponse As Integer
    ShellResponse = ShellObject.Run(ShellInstrution, 0, True) 'ShellObject.Run(Command, WindowStyle, WaitOnReturn) 0 oculta 1 Normal
    Debug.Print "shellresponse libreria: " & ShellResponse ' 0 is installed
    PythonLibraryIsInstalled = (ShellResponse = 0)
End Function

Private Sub InstallPythonLibrary()
    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    Dim Command As String
    Command = "python """ & App.localPath & "\Scripts\install_lxml.py" & """"
    Debug.Print "Comando de instalacion libreria"
    Debug.Print Command
    Dim ShellResponse As Integer
    ShellResponse = ShellObject.Run(Command, 0, True)
End Sub

Private Sub CreateOfficeThemeBasedOnOriginalThemeAndCallItCustomThemeUsingPython()

    Dim ShellObject As Object
    Set ShellObject = CreateObject("WScript.Shell")
    
    Dim ThemeFilePathInTemp As String
    ThemeFilePathInTemp = App.WorkingEnvironmentPath & ORIGINAL_THEME_FILE_NAME
    
    If Dir(ThemeFilePathInTemp) = "" Then
        MsgBox "Theme file not found: " & ThemeFilePathInTemp, vbCritical
        Exit Sub
    End If

    Dim Command As String
    Command = "python """ & App.localPath & "\Scripts\App_main.py"" """ & ThemeFilePathInTemp & """"
    
    Debug.Print "Command passed to python:"
    Debug.Print Command
    
    Dim ShellResponse As Integer
    ShellResponse = ShellObject.Run(Command, 0, True)
    
    Debug.Print "ShellResponse CreateThemeBased... " & ShellResponse

End Sub

Private Property Get TempFolderPath() As String
    TempFolderPath = Environ("TEMP") & "\"
End Property

Private Sub SaveToTemporaryFolderAsOriginalTheme()
    App.Presentation.Save
    
    Dim WorkingFolderPath As String
    WorkingFolderPath = TempFolderPath & "ToolEnvironment\"
    
    If Dir(WorkingFolderPath, vbDirectory) = "" Then
        MkDir WorkingFolderPath
    End If
    
    App.WorkingEnvironmentPath = WorkingFolderPath
    Debug.Print "Working environment path: " & WorkingFolderPath
    
    App.Presentation.SaveAs WorkingFolderPath & ORIGINAL_THEME_FILE_NAME, ppSaveAsOpenXMLTheme
    Debug.Print "Theme saved as: " & ORIGINAL_THEME_FILE_NAME
End Sub

Private Property Get LayoutDesignsSectionIndex() As Long

    LayoutDesignsSectionIndex = GetSectionIndexByName(App.Presentation, "Layout Designs")

End Property


Public Property Get LayoutDesignsFirstSlideIndex() As Long

    If LayoutDesignsSectionIndex = -1 Then Exit Property
    
    LayoutDesignsFirstSlideIndex = App.Presentation.SectionProperties.FirstSlide(LayoutDesignsSectionIndex)

End Property


Public Property Get LayoutDesignsNumberOfSlides() As Long

    If LayoutDesignsSectionIndex = -1 Then Exit Property
    
    LayoutDesignsNumberOfSlides = App.Presentation.SectionProperties.SlidesCount(LayoutDesignsSectionIndex)

End Property


Public Property Get LayoutDesignsLastSlideIndex() As Long

    If LayoutDesignsSectionIndex = -1 Then Exit Property
    
    LayoutDesignsLastSlideIndex = LayoutDesignsFirstSlideIndex + LayoutDesignsNumberOfSlides - 1

End Property
    
' // Presentation Size

Public Sub ApplySlideSizeFromMaster()

    Dim TargetSlide As Slide: Set TargetSlide = ActivePresentation.Slides(1)

    Dim WidthPixels As Double: WidthPixels = GetShapeTextAsDouble(TargetSlide, "MasterWidth")
    Dim HeightPixels As Double: HeightPixels = GetShapeTextAsDouble(TargetSlide, "MasterHeight")

    Dim WidthPoints As Double: WidthPoints = ConvertPixelsToPoints(WidthPixels)
    Dim HeightPoints As Double: HeightPoints = ConvertPixelsToPoints(HeightPixels)

    ApplySlideSize WidthPoints, HeightPoints

    Dim WidthInches As Double: WidthInches = ConvertPointsToInches(WidthPoints)
    Dim HeightInches As Double: HeightInches = ConvertPointsToInches(HeightPoints)

    Dim WidthCentimeters As Double: WidthCentimeters = ConvertInchesToCentimeters(WidthInches)
    Dim HeightCentimeters As Double: HeightCentimeters = ConvertInchesToCentimeters(HeightInches)

    Dim RatioText As String: RatioText = GetSimplifiedRatio(WidthPixels, HeightPixels)

    TargetSlide.Shapes("MasterDimensions").TextFrame.TextRange.Text = BuildDimensionsText(RatioText, WidthInches, HeightInches, WidthCentimeters, HeightCentimeters)

End Sub

Public Sub UpdateMasterDimensionsFromMasterSize()
    On Error GoTo CleanFail

    Dim TargetSlide As Slide: Set TargetSlide = ActivePresentation.Slides(1)

    Dim WidthPixels As Double: WidthPixels = GetShapeTextAsDouble(TargetSlide, "MasterWidth")
    Dim HeightPixels As Double: HeightPixels = GetShapeTextAsDouble(TargetSlide, "MasterHeight")
    If HeightPixels = 0 Then Exit Sub

    Dim WidthPoints As Double: WidthPoints = ConvertPixelsToPoints(WidthPixels)
    Dim HeightPoints As Double: HeightPoints = ConvertPixelsToPoints(HeightPixels)

    Dim WidthInches As Double: WidthInches = ConvertPointsToInches(WidthPoints)
    Dim HeightInches As Double: HeightInches = ConvertPointsToInches(HeightPoints)

    Dim WidthCentimeters As Double: WidthCentimeters = ConvertInchesToCentimeters(WidthInches)
    Dim HeightCentimeters As Double: HeightCentimeters = ConvertInchesToCentimeters(HeightInches)

    Dim RatioText As String: RatioText = GetSimplifiedRatio(WidthPixels, HeightPixels)

    TargetSlide.Shapes("MasterDimensions").TextFrame.TextRange.Text = BuildDimensionsText(RatioText, WidthInches, HeightInches, WidthCentimeters, HeightCentimeters)

    Exit Sub

CleanFail:
End Sub
Private Function GetShapeTextAsDouble(ByVal TargetSlide As Slide, ByVal ShapeName As String) As Double
    GetShapeTextAsDouble = CDbl(TargetSlide.Shapes(ShapeName).TextFrame.TextRange.Text)
End Function
Private Function ConvertPixelsToPoints(ByVal Pixels As Double) As Double
    ConvertPixelsToPoints = Pixels * 72# / 96#
End Function

Private Sub ApplySlideSize(ByVal WidthPoints As Double, ByVal HeightPoints As Double)
    With ActivePresentation.PageSetup
        .SlideWidth = WidthPoints
        .SlideHeight = HeightPoints
    End With
End Sub

Private Function ConvertPointsToInches(ByVal Points As Double) As Double
    ConvertPointsToInches = Points / 72#
End Function
Private Function ConvertInchesToCentimeters(ByVal Inches As Double) As Double
    ConvertInchesToCentimeters = Inches * 2.54
End Function
Private Function BuildDimensionsText(ByVal RatioText As String, ByVal WidthInches As Double, ByVal HeightInches As Double, ByVal WidthCentimeters As Double, ByVal HeightCentimeters As Double) As String

    BuildDimensionsText = _
        "Ratio: " & RatioText & vbCrLf & _
        "Inches: " & Format(WidthInches, "0.0") & "x" & Format(HeightInches, "0.0") & vbCrLf & _
        "Centimeters: " & Format(WidthCentimeters, "0") & "x" & Format(HeightCentimeters, "0")

End Function
Private Function GetSimplifiedRatio(ByVal Width As Double, ByVal Height As Double) As String

    Dim A As Long: A = CLng(Width)
    Dim B As Long: B = CLng(Height)

    Do While B <> 0
        Dim Temp As Long: Temp = A Mod B
        A = B
        B = Temp
    Loop

    GetSimplifiedRatio = (Width / A) & ":" & (Height / A)

End Function


'// Language

Private Sub ApplyStoredLanguageToSlideMaster()

    Dim LanguageId As MsoLanguageID: LanguageId = GetStoredLanguageId

    Dim Master As Master: Set Master = ActivePresentation.SlideMaster

    ApplyLanguageToShapes Master.Shapes, LanguageId

    Dim Layout As CustomLayout
    For Each Layout In Master.CustomLayouts
        ApplyLanguageToShapes Layout.Shapes, LanguageId
    Next Layout

End Sub
Private Sub ApplyLanguageToShapes(ByVal Shapes As Shapes, ByVal LanguageId As MsoLanguageID)

    Dim CurrentShape As Shape
    For Each CurrentShape In Shapes
        ApplyLanguageToShapeText CurrentShape, LanguageId
    Next CurrentShape

End Sub
Private Sub ApplyLanguageToShapeText(ByVal TargetShape As Shape, ByVal LanguageId As MsoLanguageID)

    If TargetShape.HasTextFrame Then
        If TargetShape.TextFrame.HasText Then
            TargetShape.TextFrame.TextRange.LanguageId = LanguageId
        End If
    End If

End Sub
Private Function GetStoredLanguageId() As MsoLanguageID

    Dim StoredValue As String
    StoredValue = ActivePresentation.Slides(1).Shapes("LanguageTextbox").AlternativeText

    If Len(StoredValue) = 0 Then
        StoredValue = CStr(ActivePresentation.Slides(1).Shapes("LanguageTextbox").TextFrame.TextRange.LanguageId)
    End If

    GetStoredLanguageId = CLng(StoredValue)

End Function


