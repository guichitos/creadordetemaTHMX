VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const IS_IN_DESIGN As Boolean = True
Private Const POINTS_PER_INCH As Double = 72#
Private Const CopyTransform As Integer = 0
Private Const GrayScaleTransform As Integer = 1
Private Const InvertLightnessTransform As Integer = 2
Private Const InvertAndGrayScaleTransform As Integer = 3

Const DefaultHeading = "Nunc viverra imperdiet enim"
Const DefaultBody = "Lorem ipsum dolor sit amet, consectetur adipiscing elit."

Private WithEvents HostApp As PowerPoint.Application
Attribute HostApp.VB_VarHelpID = -1

Private User As New AppUser

Public ThemePreviewer As New AppThemePreviewer

Public ThemeCreator As New AppThemeCreator
Public ThemeApplier As New AppThemeApplier 'TODO Try to make it Private

Private LastSelectedShapeInFontCombinations As PowerPoint.Shape

Private SavedThemeFileNameInTemp As String

Private TempMarkForFileNameTemp As String 'ya no se usa depurar

Private m_WorkingEnvironmentPath As String

Private m_IsWaitingForLayoutClick As Boolean
Private m_CurrentSlideIndex As Integer
Private m_IsFirstBackgroundCopyExecution As Boolean
Private m_HasPendingWork As Boolean
Private m_ShouldIgnoreComplexBackgrounds As Boolean
Private m_PreviousSingleSelection As PowerPoint.Shape


Friend Property Get Self() As App
    Static IsEnsuringIntegrity As Boolean
    If Not IsEnsuringIntegrity Then
        IsEnsuringIntegrity = True

        EnsureWorkingVariablesAreSet

        IsEnsuringIntegrity = False
    End If

    Set Self = Me

End Property

Private Sub EnsureWorkingVariablesAreSet()

End Sub
Public Property Get ActivePresentationInstance() As Presentation
    Set ActivePresentationInstance = ActivePresentation
End Property


Private Sub •000PastFeatures():                                                                             End Sub

'Permite usarse un tipo de letra que no está en Office solo con escribir los nombres de las fuentes en los cuadros de texto adecuados

'Make it create a OOXML theme and NOTHING else
'The created theme should
'       - Not have anything related to PowerPoint
'       - Have fmtScheme element                                Make sure the theme's lines, fills, effects, and backgrounds that appear in the UI are really useful to the user
'       - Have extraClrSchemeLst                                the option to create extra color schemes
'       - Have custClrLst                                       to give the user more options

'Allow the user to define shape effects (line, fill, effect, background) that appear in the UI

'Allow to have multiple color palettes into the same theme file

'Update Thumbnails preview ver NewTheme v3 folder
'Permitir al usuario incluir el logo de la marca para que la miniatura de tema en Office lo muestre -The BrandLogoShape is not visible now but it is there!

'Make sure the first column in the shapes gallery is more useful than black and white... Maybe remap colors using the OOXML? ClrMap

'Los efectos que hacen que las formas cambien de color deberían cambiarse

'Permitir nombrar los colores personalizados

'Los colores de hipervínculo e hipervínculo visitado no están bien en las paletas de colores

'La paleta principal del tema se puede perder al aplicar una paleta alternativa
'Crear la paleta clara, la paleta clara en escala de grises, paleta oscura y paleta oscura escala de grises automáticamente y conceptuar otras líneas de negocio con otro tema aparte
'Meter a coolors.co (abrir coolors.com/color1-color2-color3)?

Private Sub •001PresentFeatures():                                                                          End Sub



Private Sub UpdateThumbnailsPreview()
    With App.Slide
        .Shapes("16:9 Frame").BackgroundStyle = .Master.BackgroundStyle
        .Shapes("4:3 Frame").BackgroundStyle = .Master.BackgroundStyle
    End With
End Sub

Private Sub •002FutureFeatures():                                                                             End Sub

'Hacer que las variaciones de tono del mismo color para escoger no sean tan pocas, normalmente hay muchos blancos y muchos negros
'Hacer que en la parte de resultado del tema se vean los nombres de las fuentes escogidas del tema


'Calcular la paleta de colores de Office a partir de los colores principales de la marca junto con los colores neutros (blancos, grises, negros) usados por la marca
'para que el usuario de la herramienta tenga que decidir menos cosas
'Por ejemplo, para crear el tema para una marca como Coca-cola solo se debería escoger... rojo... blanco, negro y posiblemente gris
'y para crear el tema para una marca como Grada solo se debería escoger... Violeta, Amarillo... blanco, negro y posiblemente gris
'y para crear el tema para una marca como Uniformar solo se debería escoger... Naranja, Verde, Turquesa.... blanco, negro y posiblemente gris
'y se debe tomar en cuenta que para algunas marcas... el color mas oscuro a usar no es negro sino casi negro
'es posible que para otra marcas el color mas claro a usar tmapoco sea blanco sino casi blanco
'Pero la idea es que la herramienta calcule el resto de colores para crear una paleta de Office óptima

'ClrMap va a ser necesario para todas las paletas alternativas
'Para una paleta oscura
'<a:clrMap bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
'Para una paleta clara
'<a:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
'Habría que implemetar el mapeo en cada Texto alternativo de la forma que representa el color
'Por tanto seria also como así:
'<a:clrMap bg1="PaletteColorA#.AltText" tx1="PaletteColorB#.AltText" bg2="PaletteColorC#.AltText" tx2="PaletteColorD#.AltText" accent1="PaletteColorE#.AltText"....

'Y si se usa un HEX para ponerle nombre a los colores personalizados?

'Allow the user to define the 9 text styles for text boxes or else create the theme with a useful set of text styles for each layout
'Los 9 estilos pertenecen a cada textbox y no al tema o diseño o master

'Allow the user to choose the size of the slide that goes into the theme
'Que el tema sirva para 4:3, 16:9 creando un super tema a partir de dos temas _
    que tienen la particularidad de que entre ellos solo difieren por la configuración _
    de la relación de aspecto de la presentación base (para uno es 4:3 y para el otro 16:9). _
    Todo esto es útil solo para cuando se va a usar el super tema para crear una presentación en _
    PowerPoint porque para las demás aplicaciones, tener distintas relaciones de aspecto no influye o modifica nada.
    
'Set default table if possible--- No se puede


'Que si se usa el tema creado para crear una presentación haciendo click secundario y "Nuevo" entonces _
    los custom Layouts sean tan útiles y simples como los de Canva _
    y no tan inútiles y complejos como los típicos de PowerPoint. _
    https://www.canva.com/design/DAFWn9gcI1M/MLSiYaGERBbk-D1ua92bhQ/edit
'No, esto lo debe hacer la plantilla... No el tema... La plantilla idealmente debe ser basada en el tema... Pero esos detalles deben pertenecer a la plantilla

Private Sub Class_Initialize()
    Set HostApp = PowerPoint.Application
End Sub

Private Sub HostApp_SlideSelectionChanged(ByVal SldRange As SlideRange)
    RefreshRibbon
End Sub

Private Sub HostApp_WindowBeforeRightClick(ByVal Sel As Selection, Cancel As Boolean)
     If IsWaitingForLayoutClick Then
        Cancel = True
        HandleLayoutClick Sel
    End If
End Sub

Private Sub HostApp_WindowSelectionChange(ByVal Sel As Selection)
    Dim PreviousSelection As PowerPoint.Selection
    Set PreviousSelection = User.CurrentSelection

    User.AddSelection Sel


    MakeLanguageSectionDecision
    MakeFontsSectionDecisions
    MakeColorSectionDecisions
    MakeEffectSectionDecisions
    MakeResultSectionDecisions

    ResizeSlideProportionSample
    
End Sub

Private Sub ResizeSlideProportionSample()
    On Error GoTo CleanFail

    Dim WidthPixels As Double: WidthPixels = CDbl(Slide.Shapes("MasterWidth").TextFrame.TextRange.Text)
    Dim HeightPixels As Double: HeightPixels = CDbl(Slide.Shapes("MasterHeight").TextFrame.TextRange.Text)
    If HeightPixels = 0 Then Exit Sub

    Dim TargetShape As PowerPoint.Shape
    Set TargetShape = Slide.Shapes("SlideProportionSample")

    Dim FixedLeft As Double: FixedLeft = TargetShape.Left
    Dim FixedTop As Double: FixedTop = TargetShape.Top

    TargetShape.LockAspectRatio = msoFalse
    TargetShape.Height = POINTS_PER_INCH
    TargetShape.Width = POINTS_PER_INCH * (WidthPixels / HeightPixels)
    TargetShape.Left = FixedLeft
    TargetShape.Top = FixedTop

    ThemePreviewer.UpdateMasterDimensionsFromMasterSize
    
    Exit Sub

CleanFail:
End Sub

Public Property Get ThemeFileNameInTemp() As String
    ThemeFileNameInTemp = SavedThemeFileNameInTemp
End Property

Public Property Let ThemeFileNameInTemp(ByVal value As String)
    SavedThemeFileNameInTemp = value
End Property

Public Property Get TempMarkForFileName() As String
    ThemeFileNameInTemp = TempMarkForFileNameTemp
End Property

Public Property Let TempMarkForFileName(ByVal value As String)
    TempMarkForFileNameTemp = value
End Property

Public Property Get WorkingEnvironmentPath() As String
    WorkingEnvironmentPath = m_WorkingEnvironmentPath
End Property

Public Property Let WorkingEnvironmentPath(ByVal value As String)
    m_WorkingEnvironmentPath = value
End Property

Public Property Get ShouldIgnoreComplexBackgrounds() As Boolean
    ShouldIgnoreComplexBackgrounds = m_ShouldIgnoreComplexBackgrounds
End Property
Public Property Let ShouldIgnoreComplexBackgrounds(ByVal value As Boolean)
    m_ShouldIgnoreComplexBackgrounds = value
End Property


Private Sub MakeLanguageSectionDecision()
    Select Case True
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, Slide.Shapes("LanguageTextbox"))
            If User.CurrentSelection.Type = ppSelectionText Then Slide.Shapes("LanguageTextbox").Select
            Application.CommandBars.ExecuteMso "SetLanguage"
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, UpdateLanguageTextbox)
            Application.CommandBars.ExecuteMso "SetLanguage"
    End Select
    
    App.Presentation.Slides(1).Shapes("LanguageTextbox").TextFrame.TextRange.Text = GetFullLanguageNameFromLanguageId(App.Presentation.Slides(1).Shapes("LanguageTextbox").TextFrame.TextRange.LanguageId)
    App.Presentation.Slides(1).Shapes("LanguageTextbox").AlternativeText = CStr(App.Presentation.Slides(1).Shapes("LanguageTextbox").TextFrame.TextRange.LanguageId)
    
End Sub
    
Private Sub MakeFontsSectionDecisions()
    Select Case True
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, ShowTextsExamplesButton)
            ShowTextsInFontCombinations
            
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, UseFirstFontCombinationButton)
            SetShadowsBasedOn UseFirstFontCombinationButton
            
            StoreFontNamesBasedOn UseFirstFontCombinationButton
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, UseSecondFontCombinationButton)
            SetShadowsBasedOn UseSecondFontCombinationButton
            
            StoreFontNamesBasedOn UseSecondFontCombinationButton
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, UseThirdFontCombinationButton)
            SetShadowsBasedOn UseThirdFontCombinationButton
            
            StoreFontNamesBasedOn UseThirdFontCombinationButton
            
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, FirstHeadingTextShape)
            ApplyTextToWindowsAndMacFontsList FirstHeadingTextShape.TextFrame.TextRange.Text
            
            Set LastSelectedShapeInFontCombinations = FirstHeadingTextShape
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, SecondHeadingTextShape)
            ApplyTextToWindowsAndMacFontsList SecondHeadingTextShape.TextFrame.TextRange.Text
            
            Set LastSelectedShapeInFontCombinations = SecondHeadingTextShape
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, ThirdHeadingTextShape)
            ApplyTextToWindowsAndMacFontsList ThirdHeadingTextShape.TextFrame.TextRange.Text
            
            Set LastSelectedShapeInFontCombinations = ThirdHeadingTextShape
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, FirstBodyTextShape)
            ApplyTextToWindowsAndMacFontsList FirstBodyTextShape.TextFrame.TextRange.Text
            
            Set LastSelectedShapeInFontCombinations = FirstBodyTextShape
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, SecondBodyTextShape)
            ApplyTextToWindowsAndMacFontsList SecondBodyTextShape.TextFrame.TextRange.Text
            
            Set LastSelectedShapeInFontCombinations = SecondBodyTextShape
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, ThirdBodyTextShape)
            ApplyTextToWindowsAndMacFontsList ThirdBodyTextShape.TextFrame.TextRange.Text
            
            Set LastSelectedShapeInFontCombinations = ThirdBodyTextShape
            
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, SeeFontNamesButton)
            ShowFontNamesInWindowsAndMacFontsList

        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, SearchForFontsFolderButton)
            SelectFontsFolder
            
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, WindowsAndMacFontsList) And User.CurrentSelection.Type = ppSelectionText
            DisplayNameOf Font:=User.CurrentSelection.TextRange.Font
            
            TryToApplyFontToLastSelectedShapeInFontCombinations Font:=User.CurrentSelection.TextRange.Font
    End Select
End Sub

Private Sub MakeColorSectionDecisions()
    Select Case True
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, TypicalBackgroundShape):        PowerPointUtilities.ShowBackGroundFillColorPicker
            
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, AlternativeBackgroundShape):    PowerPointUtilities.ShowBackGroundFillColorPicker
        
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, HeadingTextTextbox)
            If User.CurrentSelection.Type = ppSelectionText Then HeadingTextTextbox.Select
            PowerPointUtilities.ShowTextFillColorPicker
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, TypicalTextTextbox)
            If User.CurrentSelection.Type = ppSelectionText Then TypicalTextTextbox.Select
            PowerPointUtilities.ShowTextFillColorPicker
        
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, x1stChartColorShape):           PowerPointUtilities.ShowBackGroundFillColorPicker
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, x2ndChartColorShape):           PowerPointUtilities.ShowBackGroundFillColorPicker
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, x3rdChartColorShape):           PowerPointUtilities.ShowBackGroundFillColorPicker
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, x4thChartColorShape):           PowerPointUtilities.ShowBackGroundFillColorPicker
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, x5thChartColorShape):           PowerPointUtilities.ShowBackGroundFillColorPicker
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, x6thChartColorShape):           PowerPointUtilities.ShowBackGroundFillColorPicker

        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, HyperlinkTextBox)
            If User.CurrentSelection.Type = ppSelectionText Then HyperlinkTextBox.Select
            PowerPointUtilities.ShowTextFillColorPicker

        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, VisitedHyperlinkTextBox)
            If User.CurrentSelection.Type = ppSelectionText Then VisitedHyperlinkTextBox.Select
            PowerPointUtilities.ShowTextFillColorPicker
            
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, PickedColorShape):              PowerPointUtilities.ShowBackGroundFillColorPicker

            
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, UseSuggestedAlternativePalettesButton)
            UseSuggestedAlternativePalettes
            
        
        Case IsAnAddPaletteButton(User.CurrentSelection):
            PaletteShapes(GetPaletteNumberFrom(User.CurrentSelection)).Visible = True
            Slide.Shapes("CopyCurrentColorPalette" & GetPaletteNumberFrom(User.CurrentSelection)).Visible = True
            
        Case IsARemovePaletteButton(User.CurrentSelection):
            PaletteShapes(GetPaletteNumberFrom(User.CurrentSelection)).Visible = False
            Slide.Shapes("CopyCurrentColorPalette" & GetPaletteNumberFrom(User.CurrentSelection)).Visible = False
            
        Case IsACopyCurrentColorPaletteButton(User.CurrentSelection)
            CopyCurrentColorPaletteToAlternativePalette GetPaletteNumberFrom(User.CurrentSelection)
        Case IsAnAlternativePaletteColor(User.CurrentSelection):                                        PowerPointUtilities.ShowBackGroundFillColorPicker
        
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, ShowCustomColorsButton):        SetCustomColorsVisibility True
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, HideCustomColorsButton):        SetCustomColorsVisibility False
        Case IsACustomColor(User.CurrentSelection)
            PowerPointUtilities.ShowBackGroundFillColorPicker
            PowerPointUtilities.ShowAltText
        Case User.CurrentSelection.Type = ppSelectionNone
            PutColorsInPickedColors
            UpdateBrandColorsPalette
        '/*
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, RestoreColorsForCustomColorsButton):        RestoreCustomColors
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, BlackFillForCustomColorsButton):            SetAllCustomColorsToBlack
        
        '/*
    End Select

End Sub

Private Sub CopyCurrentColorPaletteToAlternativePalette(ByVal PaletteNumber As Integer)
    SetAlternativePalette PaletteNumber, BrandColorPalette
End Sub

Private Sub UpdateBrandColorsPalette()
    PaletteShapes(0)(1).Fill.ForeColor = TypicalBackgroundShape.Fill.ForeColor
    PaletteShapes(0)(2).Fill.ForeColor = HeadingTextTextbox.TextFrame.TextRange.Font.Color
    PaletteShapes(0)(3).Fill.ForeColor = AlternativeBackgroundShape.Fill.ForeColor
    PaletteShapes(0)(4).Fill.ForeColor = TypicalTextTextbox.TextFrame.TextRange.Font.Color
    
    PaletteShapes(0)(5).Fill.ForeColor = x1stChartColorShape.Fill.ForeColor
    PaletteShapes(0)(6).Fill.ForeColor = x2ndChartColorShape.Fill.ForeColor
    PaletteShapes(0)(7).Fill.ForeColor = x3rdChartColorShape.Fill.ForeColor
    PaletteShapes(0)(8).Fill.ForeColor = x4thChartColorShape.Fill.ForeColor
    PaletteShapes(0)(9).Fill.ForeColor = x5thChartColorShape.Fill.ForeColor
    PaletteShapes(0)(10).Fill.ForeColor = x6thChartColorShape.Fill.ForeColor
    
    PaletteShapes(0)(11).Fill.ForeColor = HyperlinkTextBox.TextFrame.TextRange.Font.Color
    PaletteShapes(0)(12).Fill.ForeColor = VisitedHyperlinkTextBox.TextFrame.TextRange.Font.Color

End Sub

Private Sub PutColorsInPickedColors()
    With Slide
        Dim PickedColor As Color
        Set PickedColor = LongColor(PickedColorShape.Fill.ForeColor.Rgb)
        
        Dim i As Long
        For i = 1 To 9
            .Shapes("PickedColor+" & i).Fill.ForeColor.Rgb = PickedColor.ChangeLightnessBy(i / 18).AsLong
            .Shapes("PickedColor-" & i).Fill.ForeColor.Rgb = PickedColor.ChangeLightnessBy(-i / 18).AsLong
        Next i
    End With
End Sub

Private Sub UseSuggestedAlternativePalettes()
    With BrandColorPalette
        SetAlternativePalette 1, OfficeUtilities.GetColorPalette(.Name, _
                                                                    LongColor(.Light1).AsLong, LongColor(.Dark1).AsLong, _
                                                                    LongColor(.Light2).AsLong, LongColor(.Dark2).AsLong, _
                                                                    LongColor(.Accent1).AsLong, LongColor(.Accent2).AsLong, _
                                                                    LongColor(.Accent3).AsLong, LongColor(.Accent4).AsLong, _
                                                                    LongColor(.Accent5).AsLong, LongColor(.Accent6).AsLong, _
                                                                    LongColor(.Hyperlink).AsLong, LongColor(.VisitedHyperlink).AsLong, _
                                                                    .Mapping)
        SetAlternativePalette 2, OfficeUtilities.GetColorPalette(.Name & " - gray scale", _
                                                                    LongColor(.Light1).ToGray.AsLong, LongColor(.Dark1).ToGray.AsLong, _
                                                                    LongColor(.Light2).ToGray.AsLong, LongColor(.Dark2).ToGray.AsLong, _
                                                                    LongColor(.Accent1).ToGray.AsLong, LongColor(.Accent2).ToGray.AsLong, _
                                                                    LongColor(.Accent3).ToGray.AsLong, LongColor(.Accent4).ToGray.AsLong, _
                                                                    LongColor(.Accent5).ToGray.AsLong, LongColor(.Accent6).ToGray.AsLong, _
                                                                    LongColor(.Hyperlink).ToGray.AsLong, LongColor(.VisitedHyperlink).ToGray.AsLong, _
                                                                    .Mapping)
        Dim d As Double: d = 0
If False Then
        Dim DefaultMapping As ColorMapping
        DefaultMapping = .Mapping
        
        Dim InvertedMapping As ColorMapping
        InvertedMapping = GetMapping(.Mapping.Text1, .Mapping.Background1, .Mapping.Text2, .Mapping.Background2, .Mapping.Accent1, .Mapping.Accent2, .Mapping.Accent3, .Mapping.Accent4, .Mapping.Accent5, .Mapping.Accent6, .Mapping.Hiperlink, .Mapping.VisitedHyperlink)

        Dim Mapping2 As ColorMapping
        Mapping2 = GetMapping(.Mapping.Background2, .Mapping.Text2, .Mapping.Background1, .Mapping.Text1, .Mapping.Accent1, .Mapping.Accent2, .Mapping.Accent3, .Mapping.Accent4, .Mapping.Accent5, .Mapping.Accent6, .Mapping.Hiperlink, .Mapping.VisitedHyperlink)

        Dim InvertedMapping2 As ColorMapping
        Mapping2 = GetMapping(.Mapping.Text2, .Mapping.Background2, .Mapping.Text1, .Mapping.Background1, .Mapping.Accent1, .Mapping.Accent2, .Mapping.Accent3, .Mapping.Accent4, .Mapping.Accent5, .Mapping.Accent6, .Mapping.Hiperlink, .Mapping.VisitedHyperlink)
        
        Dim Light1InvertedDark1InvertedMappingDefault As OfficeColorPalette
        Light1InvertedDark1InvertedMappingDefault = OfficeUtilities.GetColorPalette("Light1InvDark1InvMappingDefault", _
                                                                    LongColor(.Light1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Dark1).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Light2).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Dark2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent3).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent4).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent5).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent6).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Hyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.VisitedHyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    DefaultMapping)
                                                                    
        Dim Light1InvertedDark1InvertedMappingInverted As OfficeColorPalette
        Light1InvertedDark1InvertedMappingInverted = OfficeUtilities.GetColorPalette("Light1InvDark1InvMappingInverted", _
                                                                    LongColor(.Light1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Dark1).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Light2).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Dark2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent3).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent4).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent5).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent6).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Hyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.VisitedHyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    InvertedMapping)
                                                                    
        Dim Dark1InvertedLight1InvertedMappingDefault As OfficeColorPalette
        Dark1InvertedLight1InvertedMappingDefault = OfficeUtilities.GetColorPalette("Dark1InvLight1InvMappingDefault", _
                                                                    LongColor(.Dark1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Light1).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Dark2).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Light2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent3).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent4).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent5).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent6).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Hyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.VisitedHyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    DefaultMapping)

        Dim Dark1InvertedLight1InvertedMappingInverted As OfficeColorPalette
        Dark1InvertedLight1InvertedMappingInverted = OfficeUtilities.GetColorPalette("Dark1InvLight1InvMappingInv", _
                                                                    LongColor(.Dark1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Light1).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Dark2).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Light2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent3).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent4).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent5).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent6).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Hyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.VisitedHyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    InvertedMapping)
                                                                    
        Dim Dark2InvertedLight2InvertedMappingInverted As OfficeColorPalette
        Dark2InvertedLight2InvertedMappingInverted = OfficeUtilities.GetColorPalette("Dark2InvLight2InvMappingInv", _
                                                                    LongColor(.Dark2).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Light2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Dark1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Light1).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent3).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent4).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent5).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent6).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Hyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.VisitedHyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    Mapping2)
                                                                    
        
        
        SetAlternativePalette 2, Dark1InvertedLight1InvertedMappingDefault  'Light1 is lighter than Dark1 in theme colors popup     What should be dark is light
        SetAlternativePalette 3, Light1InvertedDark1InvertedMappingInverted 'Light1 is darker than Dark1 in theme colors popup      What should be dark is light
        
        
        SetAlternativePalette 2, Light1InvertedDark1InvertedMappingDefault  'Light1 is darker than Dark1 in theme colors popup      What should be dark is dark
        SetAlternativePalette 3, Dark1InvertedLight1InvertedMappingInverted 'Light1 is lighter than Dark1 in theme colors popup     What should be dark is dark
        SetAlternativePalette 4, Dark2InvertedLight2InvertedMappingInverted
Else
        SetAlternativePalette 3, OfficeUtilities.GetColorPalette(.Name & " - inverted", _
                                                                    LongColor(.Light1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Dark1).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Light2).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Dark2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent1).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent2).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent3).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent4).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Accent5).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.Accent6).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    LongColor(.Hyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, LongColor(.VisitedHyperlink).InvertLightness.ChangeLightnessBy(d).AsLong, _
                                                                    .Mapping)
        SetAlternativePalette 4, OfficeUtilities.GetColorPalette(.Name & " - inverted gray scale", _
                                                                    LongColor(.Light1).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, LongColor(.Dark1).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, _
                                                                    LongColor(.Light2).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, LongColor(.Dark2).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, _
                                                                    LongColor(.Accent1).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, LongColor(.Accent2).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, _
                                                                    LongColor(.Accent3).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, LongColor(.Accent4).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, _
                                                                    LongColor(.Accent5).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, LongColor(.Accent6).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, _
                                                                    LongColor(.Hyperlink).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, LongColor(.VisitedHyperlink).InvertLightness.ChangeLightnessBy(d).ToGray.AsLong, _
                                                                    .Mapping)
End If
    End With
End Sub

Private Function TransformExtraColor(ByVal BaseColor As Long, ByVal Transform As Integer) As Long
    Select Case Transform
        Case CopyTransform
            TransformExtraColor = LongColor(BaseColor).AsLong
        Case GrayScaleTransform
            TransformExtraColor = LongColor(BaseColor).ToGray.AsLong
        Case InvertLightnessTransform
            TransformExtraColor = LongColor(BaseColor).InvertLightness.AsLong
        Case InvertAndGrayScaleTransform
            TransformExtraColor = LongColor(BaseColor).InvertLightness.ToGray.AsLong
    End Select
End Function

Private Function GetTransformForPalette(ByVal PaletteNumber As Integer) As Integer
    Select Case PaletteNumber
        Case 1: GetTransformForPalette = CopyTransform
        Case 2: GetTransformForPalette = GrayScaleTransform
        Case 3: GetTransformForPalette = InvertLightnessTransform
        Case 4: GetTransformForPalette = InvertAndGrayScaleTransform
        Case Else: GetTransformForPalette = CopyTransform
    End Select
End Function

Private Sub MakeEffectSectionDecisions()
    Static SelectedBySoftware As Boolean
    
    Select Case True
        Case User.CurrentSelection.Type = ppSelectionNone
            ShowAllShapesInEffectsGalleries
        Case User.CurrentSelection.Type = ppSelectionShapes, User.CurrentSelection.Type = ppSelectionText
            Dim SelectedShape As PowerPoint.Shape
            Set SelectedShape = User.CurrentSelection.ShapeRange(1)
            
            Select Case True
                Case AppUtilities.IsInArray(SelectedShape.Name, ArrayOfEffectsShapesNames) And Not SelectedBySoftware
                    ShowOnlyRowGalleriesAffectedBy SelectedShape
                Case SelectedShape.Name Like "Sample[A-G][1-9]"
                    Dim SelectedShapeInGallery As PowerPoint.Shape
                    Set SelectedShapeInGallery = SelectedShape
                
                    SelectedBySoftware = True
                    AlsoSelectTheSourceShapesOf SelectedShapeInGallery
                    SelectedBySoftware = False
            End Select
    End Select
End Sub
Private Sub ShowOnlyRowGalleriesAffectedBy(ByRef Shape As PowerPoint.Shape)
    ShowAllShapesInEffectsGalleries

    Dim EffectsNumbersVector As Variant
    EffectsNumbersVector = GetEffectNumbersThatAffectRowsInGalleryRegarding(GetEffectTypeFromShape(Shape))

    Dim ShapeSuffix As Integer
    ShapeSuffix = GetSuffixNumberOf(Shape)

    Dim i As Integer
    For i = LBound(EffectsNumbersVector) To UBound(EffectsNumbersVector)
        If EffectsNumbersVector(i) <> ShapeSuffix Then
            ShowShapesInRow i, False
        End If
    Next i
End Sub

Private Function GetSuffixNumberOf(ByRef Shape As PowerPoint.Shape) As Integer
    Dim Index As Integer
    For Index = Len(Shape.Name) To 1 Step -1
        If IsNumeric(Mid(Shape.Name, Index, 1)) Then
            GetSuffixNumberOf = Mid(Shape.Name, Index, 1)
            Exit Function
        End If
    Next Index
End Function

Private Function GetEffectNumbersThatAffectRowsInGalleryRegarding(ByRef ShapeType As String) As Variant
    If ShapeType = "Unknown" Then Exit Function

    Dim TargetColumnIndex As Integer
    TargetColumnIndex = GetEffectColumnIndexByName(TableOfAffectedRowsInGalleriesDueToThemeEffects(0), ShapeType)
    
    If TargetColumnIndex = -1 Then
        GetEffectNumbersThatAffectRowsInGalleryRegarding = Array()
        Exit Function
    End If

    Dim TotalRows As Integer
    TotalRows = UBound(TableOfAffectedRowsInGalleriesDueToThemeEffects, 1)

    Dim Result As Variant
    ReDim Result(1 To TotalRows)

    Dim MatrixRowIndex As Integer
    For MatrixRowIndex = 1 To TotalRows
        Result(MatrixRowIndex) = TableOfAffectedRowsInGalleriesDueToThemeEffects(MatrixRowIndex)(TargetColumnIndex)
    Next MatrixRowIndex

    GetEffectNumbersThatAffectRowsInGalleryRegarding = Result
End Function

Private Function GetEffectColumnIndexByName(ByVal headerRow As Variant, ByVal columnName As String) As Integer
    Dim i As Integer
    For i = LBound(headerRow, 1) To UBound(headerRow, 1)
        If headerRow(i) = columnName Then
            GetEffectColumnIndexByName = i
            Exit Function
        End If
    Next i

    GetEffectColumnIndexByName = -1
End Function


Private Function GetPaletteNumberFrom(ByRef Selection As PowerPoint.Selection) As Integer
    Dim SelectedShape As PowerPoint.Shape
    Set SelectedShape = Selection.ShapeRange(1)
    
    Select Case True
        Case IsAnAddPaletteButton(Selection):               GetPaletteNumberFrom = Replace(SelectedShape.Name, "AddPaletteButton", "")
        Case IsARemovePaletteButton(Selection):             GetPaletteNumberFrom = Replace(SelectedShape.Name, "RemovePaletteButton", "")
        Case IsACopyCurrentColorPaletteButton(Selection):   GetPaletteNumberFrom = Right(SelectedShape.Name, 1)
    End Select
End Function

Private Function IsAnAddPaletteButton(ByRef Selection As PowerPoint.Selection) As Boolean
    On Error Resume Next
    IsAnAddPaletteButton = (Selection.ShapeRange(1).Name Like "AddPaletteButton[1-4]")
    On Error GoTo 0
End Function

Private Function IsARemovePaletteButton(ByRef Selection As PowerPoint.Selection) As Boolean
    On Error Resume Next
    IsARemovePaletteButton = (Selection.ShapeRange(1).Name Like "RemovePaletteButton[1-4]")
    On Error GoTo 0
End Function

Private Function IsACopyCurrentColorPaletteButton(ByRef Selection As PowerPoint.Selection) As Boolean
    On Error Resume Next
    IsACopyCurrentColorPaletteButton = (Selection.ShapeRange(1).Name Like "CopyCurrentColorPalette[1-4]")
    On Error GoTo 0
End Function

Private Function IsAnAlternativePaletteColor(ByRef Selection As PowerPoint.Selection) As Boolean
    On Error Resume Next
    IsAnAlternativePaletteColor = (Selection.ShapeRange(1).Name Like "PaletteColor[A-R][1-4]")
    On Error GoTo 0
End Function

Private Sub MakeResultSectionDecisions()
    Select Case True
        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, ShowBrandLogoButton):           LogoShape.Visible = True

        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, HideBrandLogoButton):           LogoShape.Visible = False
        
       ' Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, SeePreviewButton):              ThemePreviewer.PreviewTheme
        
'        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, CreateThemeButton)
'            User.CurrentSelection.Unselect 'Otherwise it behaves odd
'            ThemeCreator.CreateTheme True
'        Case PowerPointUtilities.AreTheSameShape(User.CurrentSelection, ApplyThemeButton)
'            User.CurrentSelection.Unselect
'            ThemeApplier.ApplyTheme
    End Select
End Sub

Friend Property Get HeadingSampleTextbox() As PowerPoint.Shape
    Set HeadingSampleTextbox = Slide.Shapes("HeadingSampleTextbox")
End Property
Friend Property Get BodySampleTextbox() As PowerPoint.Shape
    Set BodySampleTextbox = Slide.Shapes("BodySampleTextbox")
End Property

Private Sub •01PublicProcedures():                                                                              End Sub


Private Sub •02FriendProcedures():                                                                              End Sub

Friend Property Get PaletteShapes(ByVal Number As Integer) As PowerPoint.ShapeRange
    Set PaletteShapes = Slide.Shapes.Range((Array("PaletteColorA" & Number, "PaletteColorB" & Number, "PaletteColorC" & Number, _
                                                        "PaletteColorD" & Number, "PaletteColorE" & Number, "PaletteColorF" & Number, _
                                                        "PaletteColorG" & Number, "PaletteColorH" & Number, "PaletteColorI" & Number, _
                                                        "PaletteColorJ" & Number, "PaletteColorK" & Number, "PaletteColorL" & Number)))
          ' PaletteShapes.Select
    
End Property



Friend Sub SetAlternativePalette(ByVal PaletteNumber As Integer, ByRef Colors As OfficeUtilities.OfficeColorPalette)
    PaletteNameShape(PaletteNumber).TextFrame.TextRange.Text = Colors.Name
    
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(1), Colors.Light1, ColorPlacementToString(Colors.Mapping.Background1)
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(2), Colors.Dark1, ColorPlacementToString(Colors.Mapping.Text1)
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(3), Colors.Light2, ColorPlacementToString(Colors.Mapping.Background2)
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(4), Colors.Dark2, ColorPlacementToString(Colors.Mapping.Text2)
    
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(5), Colors.Accent1, ColorPlacementToString(Colors.Mapping.Accent1)
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(6), Colors.Accent2, ColorPlacementToString(Colors.Mapping.Accent2)
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(7), Colors.Accent3, ColorPlacementToString(Colors.Mapping.Accent3)
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(8), Colors.Accent4, ColorPlacementToString(Colors.Mapping.Accent4)
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(9), Colors.Accent5, ColorPlacementToString(Colors.Mapping.Accent5)
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(10), Colors.Accent6, ColorPlacementToString(Colors.Mapping.Accent6)
    
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(11), Colors.Hyperlink, ColorPlacementToString(Colors.Mapping.Hiperlink)
    SetColorAndAltTextTo PaletteShapes(PaletteNumber)(12), Colors.VisitedHyperlink, ColorPlacementToString(Colors.Mapping.VisitedHyperlink)
End Sub
Private Sub SetColorAndAltTextTo(Shape As PowerPoint.Shape, Color As Long, AltText As String)
    With Shape
        .Fill.ForeColor.Rgb = Color
        .AlternativeText = AltText
    End With
End Sub

Friend Property Get IsInDesign() As Boolean
    IsInDesign = IS_IN_DESIGN
End Property

Friend Property Get Host() As PowerPoint.Application
    Set Host = PowerPoint.Application
End Property

Friend Property Get Presentation() As PowerPoint.Presentation
    Set Presentation = Slide.Parent
End Property

Friend Property Get Slide() As PowerPoint.Slide
    Set Slide = Slide56
End Property

Friend Property Get localPath() As String
    localPath = FileUtilities.GetLocalPath(Presentation.Path)
End Property

Friend Property Get CustomColorShapes() As PowerPoint.ShapeRange
    Dim Result As PowerPoint.ShapeRange
    Dim names() As String
    
    ReDim names(0)
    Dim i As Integer, j As Integer
    For i = 1 To 10
        For j = 1 To 5
            Dim ShapeName As String
            ShapeName = "CustomColor" & Chr(64 + i) & j
            
            ReDim Preserve names(UBound(names) + 1)
            names(UBound(names)) = ShapeName
        Next j
    Next i
    
    Set Result = Slide.Shapes.Range(names)
    
    Set CustomColorShapes = Result
End Property

Friend Property Get DefaultThemeLine() As PowerPoint.Shape
    Set DefaultThemeLine = App.Slide.Shapes("DefaultThemeLine")
    
    Debug.Assert DefaultThemeLine.Type = msoLine
End Property

Friend Property Get DefaultThemeShape() As PowerPoint.Shape
    Set DefaultThemeShape = App.DesignSlide.Shapes("DefaultThemeShape")
    
    Debug.Assert DefaultThemeShape.Type = msoAutoShape
End Property

Friend Property Get DefaultThemeTextBox() As PowerPoint.Shape
    Set DefaultThemeTextBox = App.DesignSlide.Shapes("DefaultThemeTextBox")
    
    Debug.Assert DefaultThemeTextBox.Type = msoTextBox
End Property

Friend Property Get CreateThemeButton() As PowerPoint.Shape
'    Set CreateThemeButton = Slide.Shapes("CreateThemeButton")
End Property
Friend Property Get ApplyThemeButton() As PowerPoint.Shape
'    Set ApplyThemeButton = Slide.Shapes("ApplyThemeButton")
End Property

Friend Property Get ThemeFontsBodyFontShape() As PowerPoint.Shape
    Set ThemeFontsBodyFontShape = Slide.Shapes("ThemeFontsBodyFontShape")
End Property

Friend Property Get ThemeFontsHeadingsFontShape() As PowerPoint.Shape
    Set ThemeFontsHeadingsFontShape = Slide.Shapes("ThemeFontsHeadingsFontShape")
End Property

Friend Property Get HeadingTextTextbox() As PowerPoint.Shape
    Set HeadingTextTextbox = Slide.Shapes("HeadingTextTextbox")
End Property

Friend Property Get TypicalTextTextbox() As PowerPoint.Shape
    Set TypicalTextTextbox = Slide.Shapes("TypicalTextTextbox")
End Property

Friend Property Get HyperlinkTextBox() As PowerPoint.Shape
    Set HyperlinkTextBox = Slide.Shapes("HyperlinkTextBox")
End Property

Friend Property Get VisitedHyperlinkTextBox() As PowerPoint.Shape
    Set VisitedHyperlinkTextBox = Slide.Shapes("VisitedHyperlinkTextBox")
End Property

Friend Property Get PickedColorShape() As PowerPoint.Shape
    Set PickedColorShape = Slide.Shapes("PickedColor")
End Property

Friend Property Get TypicalBackgroundShape() As PowerPoint.Shape
    Set TypicalBackgroundShape = Slide.Shapes("TypicalBackgroundShape")
End Property

Friend Property Get x1stChartColorShape() As PowerPoint.Shape
    Set x1stChartColorShape = Slide.Shapes("1stChartColorShape")
End Property
Friend Property Get x2ndChartColorShape() As PowerPoint.Shape
    Set x2ndChartColorShape = Slide.Shapes("2ndChartColorShape")
End Property
Friend Property Get x3rdChartColorShape() As PowerPoint.Shape
    Set x3rdChartColorShape = Slide.Shapes("3rdChartColorShape")
End Property
Friend Property Get x4thChartColorShape() As PowerPoint.Shape
    Set x4thChartColorShape = Slide.Shapes("4thChartColorShape")
End Property
Friend Property Get x5thChartColorShape() As PowerPoint.Shape
    Set x5thChartColorShape = Slide.Shapes("5thChartColorShape")
End Property
Friend Property Get x6thChartColorShape() As PowerPoint.Shape
    Set x6thChartColorShape = Slide.Shapes("6thChartColorShape")
End Property

Friend Property Get AlternativeBackgroundShape() As PowerPoint.Shape
    Set AlternativeBackgroundShape = Slide.Shapes("AlternativeBackgroundShape")
End Property

Friend Property Get EmptyHeadingButton() As PowerPoint.Shape
    Set EmptyHeadingButton = Slide.Shapes("EmptyHeadingButton")
End Property

Friend Property Get EmptyBodyButton() As PowerPoint.Shape
    Set EmptyBodyButton = Slide.Shapes("EmptyBodyButton")
End Property

Friend Property Get FirstHeadingTextShape() As PowerPoint.Shape
    Set FirstHeadingTextShape = Slide.Shapes("FirstHeadingTextShape")
End Property

Friend Property Get FirstBodyTextShape() As PowerPoint.Shape
    Set FirstBodyTextShape = Slide.Shapes("FirstBodyTextShape")
End Property

Friend Property Get UseFirstFontCombinationButton() As PowerPoint.Shape
    Set UseFirstFontCombinationButton = Slide.Shapes("UseFirstFontCombinationButton")
End Property

Friend Property Get SecondHeadingTextShape() As PowerPoint.Shape
    Set SecondHeadingTextShape = Slide.Shapes("SecondHeadingTextShape")
End Property

Friend Property Get SecondBodyTextShape() As PowerPoint.Shape
    Set SecondBodyTextShape = Slide.Shapes("SecondBodyTextShape")
End Property

Friend Property Get UseSecondFontCombinationButton() As PowerPoint.Shape
    Set UseSecondFontCombinationButton = Slide.Shapes("UseSecondFontCombinationButton")
End Property

Friend Property Get ThirdHeadingTextShape() As PowerPoint.Shape
    Set ThirdHeadingTextShape = Slide.Shapes("ThirdHeadingTextShape")
End Property

Friend Property Get ThirdBodyTextShape() As PowerPoint.Shape
    Set ThirdBodyTextShape = Slide.Shapes("ThirdBodyTextShape")
End Property

Friend Property Get UseThirdFontCombinationButton() As PowerPoint.Shape
    Set UseThirdFontCombinationButton = Slide.Shapes("UseThirdFontCombinationButton")
End Property

Friend Property Get SeeFontNamesButton() As PowerPoint.Shape
    Set SeeFontNamesButton = Slide.Shapes("SeeFontNamesButton")
End Property

Friend Property Get WindowsAndMacFontsList() As PowerPoint.Shape
    Set WindowsAndMacFontsList = Slide.Shapes("WindowsAndMacFontsList")
End Property

Friend Property Get SelectedFontNameShape() As PowerPoint.Shape
    Set SelectedFontNameShape = Slide.Shapes("SelectedFontNameShape")
End Property

Friend Property Get FontFolderTextbox() As PowerPoint.Shape
    Set FontFolderTextbox = Slide.Shapes("FontFolder")
End Property

Private Property Get SearchForFontsFolderButton() As PowerPoint.Shape
    Set SearchForFontsFolderButton = Slide.Shapes("SearchForFontsFolder")
End Property

Friend Property Get HeadingFontReferencePicture() As PowerPoint.Shape
    Set HeadingFontReferencePicture = Slide.Shapes("HeadingFontReferencePicture")
End Property

Friend Property Get BodyFontReferencePicture() As PowerPoint.Shape
    Set BodyFontReferencePicture = Slide.Shapes("BodyFontReferencePicture")
End Property

Friend Property Get ShowBrandLogoButton() As PowerPoint.Shape
    Set ShowBrandLogoButton = Slide.Shapes("ShowBrandLogo")
End Property

Friend Property Get LogoBrand() As PowerPoint.Shape
    Set LogoBrand = Slide.Shapes("BrandLogoImage")
End Property

Private Property Get HideBrandLogoButton() As PowerPoint.Shape
    Set HideBrandLogoButton = Slide.Shapes("HideBrandLogo")
End Property

Private Property Get ShowTextsExamplesButton() As PowerPoint.Shape
    Set ShowTextsExamplesButton = Slide.Shapes("ShowTextsExamplesButton")
End Property

Private Property Get UpdateLanguageTextbox() As PowerPoint.Shape
    Set UpdateLanguageTextbox = Slide.Shapes("LanguageTextbox")
End Property
Public Property Get DownloadsFilePath() As String
    DownloadsFilePath = Environ("USERPROFILE") & "\Downloads"
End Property
Friend Property Get ThumbnailLayout() As PowerPoint.CustomLayout

    Dim Layout As PowerPoint.CustomLayout
    
    For Each Layout In App.Presentation.SlideMaster.CustomLayouts
        If InStr(1, Layout.Name, "Thumbnail", vbTextCompare) > 0 Then
            Set ThumbnailLayout = Layout
            Exit Property
        End If
    Next Layout
    
    Set ThumbnailLayout = Nothing

End Property
Friend Property Get ThumbnailSlide() As PowerPoint.Slide
    
    If App.ActivePresentationInstance.Slides.count = 0 Then Exit Property
    
    Set ThumbnailSlide = App.ActivePresentationInstance.Slides(2)

End Property

Friend Property Get DesignSlide() As PowerPoint.Slide

    If App.ActivePresentationInstance.Slides.count = 0 Then Exit Property
    
    Set DesignSlide = App.ActivePresentationInstance.Slides(1)

End Property


Private Sub •02PrivateProcedures():                                                                              End Sub

Private Function GetEffectTypeFromShape(ByRef Shape As PowerPoint.Shape) As String
    Select Case True
        Case Shape.Name Like "Line[1-3]Style":              GetEffectTypeFromShape = "Line"
        Case Shape.Name Like "Fill[1-3]Style":              GetEffectTypeFromShape = "Fill"
        Case Shape.Name Like "Effect[1-3]Style":            GetEffectTypeFromShape = "Effect"
        Case Shape.Name Like "Background[1-3]Style":        GetEffectTypeFromShape = "Background"
        Case Else:                                          GetEffectTypeFromShape = "Unknown"
    End Select
End Function

Private Property Get UseSuggestedAlternativePalettesButton() As PowerPoint.Shape
    Set UseSuggestedAlternativePalettesButton = Slide.Shapes("UseSuggestedAlternativePalettesButton")
End Property

Private Property Get TableOfAffectedRowsInGalleriesDueToThemeEffects() As Variant
    TableOfAffectedRowsInGalleriesDueToThemeEffects = Array( _
                                                            Array("AffectedRow", "Line", "Fill", "Effect", "Background", "Text Color"), _
                                                            Array("LinesThemeGallery_Row1", 1, 0, 0, 0, 0), _
                                                            Array("LinesThemeGallery_Row2", 2, 0, 1, 0, 0), _
                                                            Array("LinesThemeGallery_Row3", 3, 0, 2, 0, 0), _
                                                            Array("ShapesThemeGallery_Row1", 2, 0, 0, 0, "dk1"), _
                                                            Array("ShapesThemeGallery_Row2", 2, 1, 0, 0, "lg1"), _
                                                            Array("ShapesThemeGallery_Row3", 3, 1, 1, 0, "lg1"), _
                                                            Array("ShapesThemeGallery_Row4", 1, 2, 1, 0, "dk1"), _
                                                            Array("ShapesThemeGallery_Row5", 1, 3, 2, 0, "lg1"), _
                                                            Array("ShapesThemeGallery_Row6", 0, 3, 3, 0, "lg1") _
                                                            )
End Property

Private Sub AlsoSelectTheSourceShapesOf(ByRef selectedGalleryShape As PowerPoint.Shape)

    Dim DependencyRowName As String
    DependencyRowName = GetRowNameFromShape(selectedGalleryShape)
    If DependencyRowName = "" Then Exit Sub

    Dim DependencyRowData As Variant
    DependencyRowData = GetDependencyRowData(DependencyRowName)
    If IsEmpty(DependencyRowData) Then Exit Sub

    Dim DependentShapeNames As Variant
    DependentShapeNames = GetDependentShapeNames(DependencyRowData)

    Dim ShapeNameIndex As Integer
    Dim SlideShape As PowerPoint.Shape
    For Each SlideShape In Slide.Shapes
        For ShapeNameIndex = LBound(DependentShapeNames) To UBound(DependentShapeNames)
            If SlideShape.Name = DependentShapeNames(ShapeNameIndex) And DependentShapeNames(ShapeNameIndex) <> "" Then
                SlideShape.Select msoFalse
            End If
        Next ShapeNameIndex
    Next SlideShape

End Sub

Private Function GetRowNameFromShape(ByRef Shape As PowerPoint.Shape) As String
    Dim rowNumber As Integer
    rowNumber = Right(Shape.Name, 1)
    
    If rowNumber >= 1 And rowNumber <= 3 Then
        GetRowNameFromShape = "LinesThemeGallery_Row" & rowNumber
    ElseIf rowNumber >= 4 And rowNumber <= 9 Then
        GetRowNameFromShape = "ShapesThemeGallery_Row" & (rowNumber - 3)
    Else
        GetRowNameFromShape = ""
    End If
End Function

Private Function GetDependencyRowData(ByVal rowName As String) As Variant
    Dim i As Integer
    For i = LBound(TableOfAffectedRowsInGalleriesDueToThemeEffects, 1) To UBound(TableOfAffectedRowsInGalleriesDueToThemeEffects, 1)
        If TableOfAffectedRowsInGalleriesDueToThemeEffects(i)(0) = rowName Then
            GetDependencyRowData = TableOfAffectedRowsInGalleriesDueToThemeEffects(i)
            Exit Function
        End If
    Next i
    GetDependencyRowData = Empty
End Function


Private Function GetDependentShapeNames(ByVal RowData As Variant) As Variant
    Dim names(0 To 2) As String
    If RowData(1) >= 1 And RowData(1) <= 3 Then names(0) = "Line" & RowData(1) & "Style"
    If RowData(2) >= 1 And RowData(2) <= 3 Then names(1) = "Fill" & RowData(2) & "Style"
    If RowData(3) >= 1 And RowData(3) <= 3 Then names(2) = "Effect" & RowData(3) & "Style"
    GetDependentShapeNames = names
End Function

Private Property Get ArrayOfEffectsShapesNames() As Variant
    ArrayOfEffectsShapesNames = Array("Line1Style", "Line2Style", "Line3Style", _
                                        "Fill1Style", "Fill2Style", "Fill3Style", _
                                        "Effect1Style", "Effect2Style", "Effect3Style")
End Property

Private Sub ShowAllShapesInEffectsGalleries()
    Dim rowIndex As Integer, columnIndex As Integer
    For rowIndex = LBound(SampleShapesMatrix, 1) To UBound(SampleShapesMatrix, 1)
        For columnIndex = LBound(SampleShapesMatrix, 2) To UBound(SampleShapesMatrix, 2)
            Dim ShapeName As String
            ShapeName = SampleShapesMatrix(rowIndex, columnIndex)

            Dim ShapeElement As Shape
            On Error Resume Next
            Set ShapeElement = App.Slide.Shapes(ShapeName)
            If Not ShapeElement Is Nothing Then
                ShapeElement.Visible = msoTrue
            End If
            On Error GoTo 0
        Next columnIndex
    Next rowIndex
End Sub

Private Property Get SampleShapesMatrix() As Variant
    Dim ShapesMatrix(1 To 9, 1 To 7) As String
    Dim rowIndex As Integer, columnIndex As Integer

    For rowIndex = 1 To 9
        For columnIndex = 1 To 7
            ShapesMatrix(rowIndex, columnIndex) = "Sample" & Chr(64 + columnIndex) & rowIndex
        Next columnIndex
    Next rowIndex

    SampleShapesMatrix = ShapesMatrix
End Property

Private Function FilterSampleShapesByRow(ShapeMatrix As Variant, TargetRow As Integer) As Variant
    Dim FilteredShapes(1 To 7) As String

    If TargetRow < 1 Or TargetRow > 9 Then
        MsgBox "Row number out of range. Must be between 1 and 9.", vbExclamation, "Error"
        Exit Function
    End If

    Dim columnIndex As Integer
    For columnIndex = 1 To 7
        FilteredShapes(columnIndex) = ShapeMatrix(TargetRow, columnIndex)
    Next columnIndex

    FilterSampleShapesByRow = FilteredShapes
End Function

Private Sub SetRowShapesVisibility(FilteredRowShapeMatrix As Variant, IsVisible As Boolean)
    Dim rowIndex As Integer
    Dim ShapeName As String
    Dim ShapeElement As Shape
    For rowIndex = LBound(FilteredRowShapeMatrix, 1) To UBound(FilteredRowShapeMatrix, 1)
        ShapeName = FilteredRowShapeMatrix(rowIndex)

        On Error Resume Next
        Set ShapeElement = App.Slide.Shapes(ShapeName)
        If Not ShapeElement Is Nothing Then
            ShapeElement.Visible = IIf(IsVisible, msoTrue, msoFalse)
        End If
        On Error GoTo 0
    Next rowIndex
End Sub

Private Sub ShowShapesInRow(TargetRow As Integer, IsVisible As Boolean)
    Dim FilteredShapes As Variant
    FilteredShapes = FilterSampleShapesByRow(SampleShapesMatrix, TargetRow)

    SetRowShapesVisibility FilteredShapes, IsVisible
End Sub

Private Property Get SeePreviewButton() As PowerPoint.Shape
   ' Set SeePreviewButton = Slide.Shapes("SeePreview")
End Property

Private Property Get BrandColorPalette() As OfficeUtilities.OfficeColorPalette
    With PaletteShapes(0)
        BrandColorPalette = OfficeUtilities.GetColorPalette("Brand", _
                                                                .Item(1).Fill.ForeColor.Rgb, .Item(2).Fill.ForeColor.Rgb, _
                                                                .Item(3).Fill.ForeColor.Rgb, .Item(4).Fill.ForeColor.Rgb, _
                                                                .Item(5).Fill.ForeColor.Rgb, .Item(6).Fill.ForeColor.Rgb, .Item(7).Fill.ForeColor.Rgb, .Item(8).Fill.ForeColor.Rgb, .Item(9).Fill.ForeColor.Rgb, .Item(10).Fill.ForeColor.Rgb, _
                                                                .Item(11).Fill.ForeColor.Rgb, .Item(12).Fill.ForeColor.Rgb, _
                                                                DefaultColorMapping)
    End With
End Property

Private Property Get PaletteNameShape(PaletteNumber As Integer) As PowerPoint.Shape
    Set PaletteNameShape = Slide.Shapes("PaletteName" & PaletteNumber)
End Property

Private Sub ShowFontNamesInWindowsAndMacFontsList()
    With App.WindowsAndMacFontsList.TextFrame.TextRange
        Dim n As Long
        For n = 1 To .Paragraphs.count + 1
            .Paragraphs(n - 1, 1).Text = .Paragraphs(n - 1, 1).Font.Name
        Next n
    End With
End Sub

Private Sub TryToApplyFontToLastSelectedShapeInFontCombinations(ByRef Font As PowerPoint.Font)
    If LastSelectedShapeInFontCombinations Is Nothing Then Exit Sub
    
    Select Case True
        Case FirstHeadingTextShape Is LastSelectedShapeInFontCombinations:  UseFirstFontCombinationButton.TextFrame.TextRange.Paragraphs(1).Text = Font.Name
        Case FirstBodyTextShape Is LastSelectedShapeInFontCombinations:     UseFirstFontCombinationButton.TextFrame.TextRange.Paragraphs(2).Text = Font.Name

        Case SecondHeadingTextShape Is LastSelectedShapeInFontCombinations: UseSecondFontCombinationButton.TextFrame.TextRange.Paragraphs(1).Text = Font.Name
        Case SecondBodyTextShape Is LastSelectedShapeInFontCombinations:    UseSecondFontCombinationButton.TextFrame.TextRange.Paragraphs(2).Text = Font.Name

        Case ThirdHeadingTextShape Is LastSelectedShapeInFontCombinations:  UseThirdFontCombinationButton.TextFrame.TextRange.Paragraphs(1).Text = Font.Name
        Case ThirdBodyTextShape Is LastSelectedShapeInFontCombinations:     UseThirdFontCombinationButton.TextFrame.TextRange.Paragraphs(2).Text = Font.Name
    End Select
    
    PowerPointUtilities.ApplyFontToShape LastSelectedShapeInFontCombinations, Font.Name
End Sub

Private Sub DisplayNameOf(ByRef Font As PowerPoint.Font)
    App.SelectedFontNameShape.TextFrame.TextRange.Text = Font.Name
End Sub

Private Sub SelectFontsFolder()
    Dim InitialFontsFolderPath As String
    InitialFontsFolderPath = App.localPath

    Dim SelectedFontsFolderPath As String
    SelectedFontsFolderPath = BrowseForFolder(initialPath:=InitialFontsFolderPath, dialogTitle:="Select fonts folder")
    If Len(SelectedFontsFolderPath) = 0 Then Exit Sub

    FontFolderTextbox.TextFrame.TextRange.Text = SelectedFontsFolderPath
End Sub

Private Sub SetShadowsBasedOn(ByRef Shape As PowerPoint.Shape)
    Select Case True
        Case Shape Is UseFirstFontCombinationButton
            AddShadowToShape UseFirstFontCombinationButton
            RemoveShadowToShape UseSecondFontCombinationButton
            RemoveShadowToShape UseThirdFontCombinationButton
        Case Shape Is UseSecondFontCombinationButton
            RemoveShadowToShape UseFirstFontCombinationButton
            AddShadowToShape UseSecondFontCombinationButton
            RemoveShadowToShape UseThirdFontCombinationButton
        Case Shape Is UseThirdFontCombinationButton
            RemoveShadowToShape UseFirstFontCombinationButton
            RemoveShadowToShape UseSecondFontCombinationButton
            AddShadowToShape UseThirdFontCombinationButton
    End Select
End Sub

Private Sub StoreFontNamesBasedOn(ByRef Shape As PowerPoint.Shape)
    Dim HeadingsFont As String
    Dim BodyFont As String
    
    Select Case True
        Case Shape Is UseFirstFontCombinationButton
            HeadingsFont = FirstHeadingTextShape.TextFrame.TextRange.Font.Name
            BodyFont = FirstBodyTextShape.TextFrame.TextRange.Font.Name
        Case Shape Is UseSecondFontCombinationButton
            HeadingsFont = SecondHeadingTextShape.TextFrame.TextRange.Font.Name
            BodyFont = SecondBodyTextShape.TextFrame.TextRange.Font.Name
        Case Shape Is UseThirdFontCombinationButton
            HeadingsFont = ThirdHeadingTextShape.TextFrame.TextRange.Font.Name
            BodyFont = ThirdBodyTextShape.TextFrame.TextRange.Font.Name
    End Select
    
    HeadingsAlternativeFontNameTextbox.TextFrame.TextRange.Text = HeadingsFont
    BodyAlternativeFontNameTextbox.TextFrame.TextRange.Text = BodyFont
End Sub

Friend Property Get HeadingsAlternativeFontNameTextbox() As PowerPoint.Shape
    Set HeadingsAlternativeFontNameTextbox = Slide.Shapes("HeadingsAlternativeFontNameTextbox")
End Property

Friend Property Get BodyAlternativeFontNameTextbox() As PowerPoint.Shape
    Set BodyAlternativeFontNameTextbox = Slide.Shapes("BodyAlternativeFontNameTextbox")
End Property

Private Property Get ShowCustomColorsButton() As PowerPoint.Shape
    Set ShowCustomColorsButton = Slide.Shapes("ShowCustomColors")
End Property

Private Property Get HideCustomColorsButton() As PowerPoint.Shape
    Set HideCustomColorsButton = Slide.Shapes("HideCustomColors")
End Property

Private Property Get RestoreColorsForCustomColorsButton() As PowerPoint.Shape
    Set RestoreColorsForCustomColorsButton = Slide.Shapes("RestoreColorsForCustomColors")
End Property
Private Property Get BlackFillForCustomColorsButton() As PowerPoint.Shape
    Set BlackFillForCustomColorsButton = Slide.Shapes("BlackFillForCustomColors")
End Property

Private Function IsACustomColor(ByRef Selection As PowerPoint.Selection) As Boolean
    Dim Result As Boolean
    
    Dim Shape As PowerPoint.Shape
    For Each Shape In CustomColorShapes
        Result = PowerPointUtilities.AreTheSameShape(Selection, Shape)
        
        If Result = True Then Exit For
    Next Shape
    
    IsACustomColor = Result
End Function

Private Property Get LogoShape() As PowerPoint.Shape
    Set LogoShape = Slide.Shapes("BrandLogoImage")
End Property

Private Sub SetCustomColorsVisibility(showColors As Boolean)
    Dim i As Long
    For i = 1 To CustomColorShapes.count
        CustomColorShapes(i).Visible = (showColors = True)
    Next i
    
    App.Slide.Shapes("CustomColorLabel").Visible = (showColors = True)
End Sub

Private Sub ApplyTextToWindowsAndMacFontsList(ByVal Text As String)
    With App.WindowsAndMacFontsList.TextFrame.TextRange
        Dim n As Long
        For n = 1 To .Paragraphs.count + 1
            .Paragraphs(n - 1, 1).Text = Left(Text, Len("WWWWW WWWWW WWWWW")) & "..." 'So texts don't split in two lines
        Next n
    End With
End Sub

Private Property Get HeadingText() As String
    HeadingText = HeadingSampleTextbox.TextFrame.TextRange.Text
End Property

Private Property Get BodyText() As String
    BodyText = BodySampleTextbox.TextFrame.TextRange.Text
End Property
Public Property Get FontsInWindowsAndMacFontsList() As Collection

    Dim FontList As New Collection

    Dim TargetSlide As Slide: Set TargetSlide = ActivePresentation.Slides(1)

    Dim FontListShape As Shape
    On Error Resume Next
    Set FontListShape = TargetSlide.Shapes("WindowsAndMacFontsList")
    On Error GoTo 0

    If FontListShape Is Nothing Then
        Set FontsInWindowsAndMacFontsList = FontList
        Exit Property
    End If

    If Not FontListShape.HasTextFrame Then
        Set FontsInWindowsAndMacFontsList = FontList
        Exit Property
    End If

    If Not FontListShape.TextFrame.HasText Then
        Set FontsInWindowsAndMacFontsList = FontList
        Exit Property
    End If

    Dim FontRegistry As Object: Set FontRegistry = CreateObject("Scripting.Dictionary")

    Dim Paragraph As TextRange
    For Each Paragraph In FontListShape.TextFrame.TextRange.Paragraphs

        Dim FontName As String: FontName = Paragraph.Characters(1, 1).Font.Name

        If Len(FontName) > 0 Then
            If Not FontRegistry.Exists(FontName) Then
                FontRegistry.Add FontName, True
                FontList.Add FontName
            End If
        End If

    Next Paragraph

    Set FontsInWindowsAndMacFontsList = FontList

End Property
Public Function IsFontInWindowsAndMacFontsList(ByVal FontNameToCheck As String) As Boolean

    Dim ExistingFontName As Variant
    For Each ExistingFontName In Me.FontsInWindowsAndMacFontsList
        If StrComp(ExistingFontName, FontNameToCheck, vbTextCompare) = 0 Then
            IsFontInWindowsAndMacFontsList = True
            Exit Function
        End If
    Next ExistingFontName

    IsFontInWindowsAndMacFontsList = False

End Function
Public Property Get BothBrandFontsAreInWindowsAndMacFontsList() As Boolean

    Dim TargetSlide As Slide: Set TargetSlide = ActivePresentation.Slides(1)

    Dim HeadingShape As Shape
    On Error Resume Next
    Set HeadingShape = TargetSlide.Shapes("HeadingsBrandFontNameTextbox")
    On Error GoTo 0
    If HeadingShape Is Nothing Then Exit Property
    If Not HeadingShape.HasTextFrame Then Exit Property
    If Not HeadingShape.TextFrame.HasText Then Exit Property

    Dim BodyShape As Shape
    On Error Resume Next
    Set BodyShape = TargetSlide.Shapes("BodyBrandFontNameTextbox")
    On Error GoTo 0
    If BodyShape Is Nothing Then Exit Property
    If Not BodyShape.HasTextFrame Then Exit Property
    If Not BodyShape.TextFrame.HasText Then Exit Property

    Dim HeadingFontName As String: HeadingFontName = Trim(HeadingShape.TextFrame.TextRange.Text)
    Dim BodyFontName As String: BodyFontName = Trim(BodyShape.TextFrame.TextRange.Text)

    If Me.IsFontInWindowsAndMacFontsList(HeadingFontName) _
       And Me.IsFontInWindowsAndMacFontsList(BodyFontName) Then

        BothBrandFontsAreInWindowsAndMacFontsList = True
    End If

End Property



Private Sub ShowTextsInFontCombinations()
    Dim DisplayedHeadingValue As String
    Select Case Len(HeadingText)
        Case 0:         DisplayedHeadingValue = String(Len(DefaultHeading), "_") 'So the user sees the shapee
        Case Else:      DisplayedHeadingValue = Trim(Left(HeadingText, Len(DefaultHeading)))
    End Select
    
    FirstHeadingTextShape.TextFrame.TextRange.Text = DisplayedHeadingValue
    SecondHeadingTextShape.TextFrame.TextRange.Text = DisplayedHeadingValue
    ThirdHeadingTextShape.TextFrame.TextRange.Text = DisplayedHeadingValue
    
    Dim DisplayedBodyValue As String
    Select Case Len(BodyText)
        Case 0:         DisplayedBodyValue = String(Len(DefaultBody), "_") 'So the user sees the shape
        Case Else:      DisplayedBodyValue = Trim(Left(BodyText, Len(DefaultBody)))
    End Select
    
    FirstBodyTextShape.TextFrame.TextRange.Text = DisplayedBodyValue
    SecondBodyTextShape.TextFrame.TextRange.Text = DisplayedBodyValue
    ThirdBodyTextShape.TextFrame.TextRange.Text = DisplayedBodyValue
End Sub

Private Sub AddShadowToShape(ByRef Shape As PowerPoint.Shape)
    With Shape.Shadow
        .Visible = msoTrue
        .ForeColor.Rgb = Rgb(0, 0, 0)
        .Transparency = 0.5
        .Blur = 4
        .OffsetX = 3
        .OffsetY = 3
    End With
End Sub

Private Sub RemoveShadowToShape(ByRef Shape As PowerPoint.Shape)
    Shape.Shadow.Visible = msoFalse
End Sub

'/*

Private Sub RestoreCustomColors()
    Dim TargetSlide As Slide: Set TargetSlide = ActivePresentation.Slides(1)
    
    Dim ShapeName As Variant

    Dim CustomShape As Shape
    Dim RowLetter As Variant
    Dim ColumnNumber As Integer
    Dim ColorRGB As Long
    
    Dim ColorMap As Object: Set ColorMap = CreateObject("Scripting.Dictionary")
    
    ColorMap.Add "CustomColorA1", Rgb(227, 227, 227)
    ColorMap.Add "CustomColorA2", Rgb(250, 221, 189)
    ColorMap.Add "CustomColorA3", Rgb(62, 103, 35)
    ColorMap.Add "CustomColorA4", Rgb(140, 197, 64)
    ColorMap.Add "CustomColorA5", Rgb(39, 39, 39)
    
    ColorMap.Add "CustomColorB1", Rgb(238, 172, 0)
    ColorMap.Add "CustomColorB2", Rgb(248, 217, 171)
    ColorMap.Add "CustomColorB3", Rgb(213, 208, 20)
    ColorMap.Add "CustomColorB4", Rgb(249, 243, 26)
    ColorMap.Add "CustomColorB5", Rgb(89, 89, 89)
    
    ColorMap.Add "CustomColorC1", Rgb(220, 91, 7)
    ColorMap.Add "CustomColorC2", Rgb(234, 192, 152)
    ColorMap.Add "CustomColorC3", Rgb(201, 97, 22)
    ColorMap.Add "CustomColorC4", Rgb(228, 111, 27)
    ColorMap.Add "CustomColorC5", Rgb(149, 149, 149)
    
    ColorMap.Add "CustomColorD1", Rgb(169, 107, 61)
    ColorMap.Add "CustomColorD2", Rgb(236, 178, 123)
    ColorMap.Add "CustomColorD3", Rgb(139, 69, 19)
    ColorMap.Add "CustomColorD4", Rgb(156, 109, 92)
    ColorMap.Add "CustomColorD5", Rgb(170, 170, 170)
    
    ColorMap.Add "CustomColorE1", Rgb(204, 92, 78)
    ColorMap.Add "CustomColorE2", Rgb(220, 166, 118)
    ColorMap.Add "CustomColorE3", Rgb(194, 19, 24)
    ColorMap.Add "CustomColorE4", Rgb(228, 57, 26)
    ColorMap.Add "CustomColorE5", Rgb(188, 188, 188)
    
    ColorMap.Add "CustomColorF1", Rgb(215, 71, 120)
    ColorMap.Add "CustomColorF2", Rgb(227, 157, 87)
    ColorMap.Add "CustomColorF3", Rgb(215, 69, 119)
    ColorMap.Add "CustomColorF4", Rgb(255, 186, 191)
    ColorMap.Add "CustomColorF5", Rgb(203, 203, 203)
    
    ColorMap.Add "CustomColorG1", Rgb(137, 88, 55)
    ColorMap.Add "CustomColorG2", Rgb(208, 130, 63)
    ColorMap.Add "CustomColorG3", Rgb(211, 0, 211)
    ColorMap.Add "CustomColorG4", Rgb(255, 124, 255)
    ColorMap.Add "CustomColorG5", Rgb(218, 218, 218)
    
    ColorMap.Add "CustomColorH1", Rgb(79, 47, 26)
    ColorMap.Add "CustomColorH2", Rgb(195, 133, 108)
    ColorMap.Add "CustomColorH3", Rgb(199, 142, 254)
    ColorMap.Add "CustomColorH4", Rgb(232, 206, 255)
    ColorMap.Add "CustomColorH5", Rgb(231, 231, 231)
    
    ColorMap.Add "CustomColorI1", Rgb(50, 24, 7)
    ColorMap.Add "CustomColorI2", Rgb(192, 123, 74)
    ColorMap.Add "CustomColorI3", Rgb(108, 180, 205)
    ColorMap.Add "CustomColorI4", Rgb(127, 211, 240)
    ColorMap.Add "CustomColorI5", Rgb(243, 243, 243)
    
    ColorMap.Add "CustomColorJ1", Rgb(51, 25, 10)
    ColorMap.Add "CustomColorJ2", Rgb(142, 88, 64)
    ColorMap.Add "CustomColorJ3", Rgb(46, 135, 182)
    ColorMap.Add "CustomColorJ4", Rgb(130, 168, 199)
    ColorMap.Add "CustomColorJ5", Rgb(253, 253, 253)

    For Each ShapeName In ColorMap.Keys
        On Error Resume Next
        Set CustomShape = TargetSlide.Shapes(ShapeName)
        On Error GoTo 0
        
        If Not CustomShape Is Nothing Then
            CustomShape.Fill.ForeColor.Rgb = ColorMap(ShapeName)
        End If
        
        Set CustomShape = Nothing
    Next ShapeName
End Sub

Private Sub SetAllCustomColorsToBlack()

    Dim FillColorRGB As Long: FillColorRGB = Rgb(0, 0, 0)
    
    Dim RowLetter As Variant
    Dim ColumnNumber As Integer
    Dim ShapeName As String
    Dim CustomShape As Shape
    
    For Each RowLetter In Split("A B C D E F G H I J")
        For ColumnNumber = 1 To 5
            ShapeName = "CustomColor" & RowLetter & ColumnNumber
            
            On Error Resume Next
            Set CustomShape = ActivePresentation.Slides(1).Shapes(ShapeName)
            On Error GoTo 0
            
            If Not CustomShape Is Nothing Then
                If CustomShape.Fill.Visible Then
                    CustomShape.Fill.ForeColor.Rgb = FillColorRGB
                End If
                Set CustomShape = Nothing
            End If
        Next ColumnNumber
    Next RowLetter

End Sub

Friend Property Get IsWaitingForLayoutClick() As Boolean
    IsWaitingForLayoutClick = m_IsWaitingForLayoutClick
End Property

Friend Property Let IsWaitingForLayoutClick(ByVal value As Boolean)
    m_IsWaitingForLayoutClick = value
End Property
Friend Property Get CurrentSlideIndex() As Integer
    CurrentSlideIndex = m_CurrentSlideIndex
End Property

Friend Property Let CurrentSlideIndex(ByVal value As Integer)
    m_CurrentSlideIndex = value
End Property

Friend Property Get IsFirstBackgroundCopyExecution() As Boolean
    IsFirstBackgroundCopyExecution = m_IsFirstBackgroundCopyExecution
End Property

Friend Property Let IsFirstBackgroundCopyExecution(ByVal value As Boolean)
    m_IsFirstBackgroundCopyExecution = value
End Property

Public Property Get HasPendingWork() As Boolean
    HasPendingWork = m_HasPendingWork
End Property

Public Property Let HasPendingWork(ByVal value As Boolean)
    m_HasPendingWork = value
End Property

Public Property Get LayoutDesignHasComplexBackgrounds() As Boolean

    Dim SectionIndex As Long: SectionIndex = GetSectionIndexByName(App.Presentation, "Layout Designs")
    If SectionIndex = 0 Then Exit Property

    Dim FirstSlideIndex As Long
    FirstSlideIndex = App.ThemePreviewer.LayoutDesignsFirstSlideIndex

    Dim SlideCount As Long
    SlideCount = App.ThemePreviewer.LayoutDesignsNumberOfSlides

    Dim SlideIndex As Long
    For SlideIndex = FirstSlideIndex To FirstSlideIndex + SlideCount - 1

        Dim CurrentSlide As Slide: Set CurrentSlide = App.Presentation.Slides(SlideIndex)

        Select Case CurrentSlide.Background.Fill.Type
            Case msoFillPicture, msoFillTextured, msoFillGradient
                LayoutDesignHasComplexBackgrounds = True
                Exit Property
        End Select

    Next SlideIndex

End Property

Private Sub HandleLayoutClick(ByVal Sel As Selection)
    Application.CommandBars.ExecuteMso "MasterViewClose"
    App.IsWaitingForLayoutClick = False
    SyncLayoutsFromSampleSlidesShapesAndPlaceholders
End Sub


