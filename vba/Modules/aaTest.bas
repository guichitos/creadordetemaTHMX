Attribute VB_Name = "aaTest"
Public Sub ApplyPreviewThemeForTesting()
   
    'Dim PreviousDesignName As String: PreviousDesignName = App.Slide.Design.Name
    Dim CustomThemePath As String: CustomThemePath = App.WorkingEnvironmentPath & "CustomTheme.thmx"
    Debug.Print CustomThemePath
    ClearNonPlaceholderShapesFromLayoutsInLayoutDesignSection 'Otherwise shapes in layouts are duplicated
    App.Presentation.ApplyTheme CustomThemePath
    App.ThemePreviewer.ResetValuesForLayoutFlow
    App.Slide.Design.Preserved = False 'Otherwise PowerPoint creates duplicates
   ' App.Slide.Design.Name = PreviousDesignName
    
    App.ThemeFontsHeadingsFontShape.TextFrame.TextRange.Text = App.HeadingsAlternativeFontNameTextbox.TextFrame.TextRange.Text
    App.ThemeFontsBodyFontShape.TextFrame.TextRange.Text = App.BodyAlternativeFontNameTextbox.TextFrame.TextRange.Text
    'RemoveBrandLogoImagesFromOriginalTheme 'Must be deleted
End Sub
