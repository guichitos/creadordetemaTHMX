VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSelections As New Collection

Friend Sub AddSelection(ByRef Selection As PowerPoint.Selection)
    pSelections.Add Selection
End Sub

Private Property Get Selections() As Collection
    Set Selections = pSelections
End Property

Friend Property Get CurrentSelection() As PowerPoint.Selection
    If Selections.count = 0 Then
        Set CurrentSelection = Nothing
    Else
        Set CurrentSelection = Selections.Item(Selections.count)
    End If
End Property




'Friend Sub WantsToPickACustomColor()
'    PowerPointUtilities.ShowBackGroundFillColorPicker
'End Sub

'Friend Sub SelectedTypicalBackgroundShape()
'    App.ShowBackGroundFillColorPicker
'End Sub

'Friend Sub SelectedOneOfTheChartColorShapes()
'    App.ShowBackGroundFillColorPicker
'End Sub

'Friend Sub SelectedAlternativeBackgroundShape()
'    App.ShowBackGroundFillColorPicker
'End Sub

'Friend Sub SelectedHeadingTextTextBox()
'    App.ShowTextFillColorPicker
'End Sub
'Friend Sub SelectedTextInHeadingTextTextBox()
'    App.HeadingTextTextbox.Select
'
'    SelectedHeadingTextTextBox
'End Sub

'Friend Sub SelectedTypicalTextTextBox()
'    App.ShowTextFillColorPicker
'End Sub
'Friend Sub SelectedTextInTypicalTextTextBox()
'    App.TypicalTextTextbox.Select
'
'    SelectedTypicalTextTextBox
'End Sub

'Friend Sub SelectedHyperlinkTextBox()
'    App.ShowTextFillColorPicker
'End Sub
'Friend Sub SelectedVisitedHyperlinkTextBox()
'
'    App.ShowTextFillColorPicker
'End Sub

'Friend Sub SelectedTextInHyperlinkTextBox()
'    App.HyperlinkTextBox.Select
'    If MsgBox("Do you want to apply the same color to visited links?", vbYesNo) = vbYes Then
'        App.Slide.Shapes.Range(Array("HyperlinkTextBox", "VisitedHyperlinkTextBox")).Select
'    End If
'
'    SelectedHyperlinkTextBox
'End Sub

'Friend Sub UserSelectedTextInVisitedHyperlinkTextBox()
'    App.VisitedHyperlinkTextBox.Select
'
'    SelectedHyperlinkTextBox
'End Sub

Friend Sub WantsToUseFonts(ByRef HeadingFont As PowerPoint.Font, ByRef BodyFont As PowerPoint.Font)
    PowerPointUtilities.ApplyFontToShape App.HeadingTextTextbox, HeadingFont
    
    PowerPointUtilities.ApplyFontToShape App.TypicalTextTextbox, BodyFont
    PowerPointUtilities.ApplyFontToShape App.HyperlinkTextBox, BodyFont
    PowerPointUtilities.ApplyFontToShape App.VisitedHyperlinkTextBox, BodyFont
End Sub

Friend Sub WantsToUseFirstFontCombination()
    WantsToUseFonts GetFontOf(App.FirstHeadingTextShape), GetFontOf(App.FirstBodyTextShape)
End Sub

Friend Sub WantsToUseSecondFontCombination()
    WantsToUseFonts GetFontOf(App.SecondHeadingTextShape), GetFontOf(App.SecondBodyTextShape)
End Sub

Friend Sub WantsToUseThirdFontCombination()
    WantsToUseFonts GetFontOf(App.ThirdHeadingTextShape), GetFontOf(App.ThirdBodyTextShape)
End Sub

