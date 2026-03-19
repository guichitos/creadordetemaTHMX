Attribute VB_Name = "AppLayoutSynchronizer"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public Sub SyncLayoutsFromSampleSlidesShapesAndPlaceholders()
    UpdateBrandLogoInThumbnailLayout
    UpdateAllBrandLogoGraphicsInSlides
    
    If App Is Nothing Then Set App = New App

    Dim presentationInstance As PowerPoint.Presentation
    Set presentationInstance = App.ActivePresentationInstance

    Dim SampleSlideIndex As Long: SampleSlideIndex = App.ThemePreviewer.LayoutDesignsFirstSlideIndex

    Do While SampleSlideIndex <= App.ThemePreviewer.LayoutDesignsLastSlideIndex
        
        If SampleSlideIndex > App.CurrentSlideIndex Then
            
            App.CurrentSlideIndex = SampleSlideIndex

            If App.CurrentSlideIndex = App.ThemePreviewer.LayoutDesignsFirstSlideIndex Then ClearNonPlaceholderShapesFromLayoutsInLayoutDesignSection
            
            Dim SampleSlide As PowerPoint.Slide
            Set SampleSlide = presentationInstance.Slides(SampleSlideIndex)

            On Error Resume Next
            'SampleSlide.Select
            On Error GoTo 0

            SyncLayoutFromSampleSlideShapes SampleSlide
            ApplySlideBackgroundToLayout SampleSlide, SampleSlide.CustomLayout

        End If

        If App.IsWaitingForLayoutClick Then Exit Sub

        SampleSlideIndex = SampleSlideIndex + 1

    Loop

    App.CurrentSlideIndex = 0

End Sub
Private Sub UpdateBrandLogoInThumbnailLayout()
    Dim Shape As PowerPoint.Shape
    Dim i As Long
    For i = App.ThumbnailLayout.Shapes.count To 1 Step -1
        Set Shape = App.ThumbnailLayout.Shapes(i)
        If Shape.Name = BrandLogoImage.Name Then
            Shape.Delete 'Sometimes PowerPoint creates duplicates
        End If
    Next i
    
    If BrandLogoImage.Visible Then
        BrandLogoImage.Copy
        
        Dim PastedShape As PowerPoint.Shape
        Set PastedShape = App.ThumbnailLayout.Shapes.Paste(1)
        
        With PastedShape
            .Width = (1 / 3) * App.ThumbnailLayout.Width
            .Left = (3 / 4) * App.ThumbnailLayout.Width - (1 / 2) * .Width
            
            .Top = 0
            Do Until .Top + (1 / 2) * .Height >= (1 / 4) * App.ThumbnailLayout.Height
                .Top = .Top + 1
            Loop
        End With
    End If
End Sub

Private Sub UpdateAllBrandLogoGraphicsInSlides()

    Dim PrimaryLogoName As String: PrimaryLogoName = BrandLogoImage.Name
    Dim PositiveMarkName As String: PositiveMarkName = "PositiveMark"
    Dim NegativeLogoName As String: NegativeLogoName = "NegativeLogo"
    Dim NegativeMarkName As String: NegativeMarkName = "NegativeMark"
    
    UpdateGraphicInSlides PrimaryLogoName
    UpdateGraphicInSlides PositiveMarkName
    UpdateGraphicInSlides NegativeLogoName
    UpdateGraphicInSlides NegativeMarkName

End Sub

Private Sub UpdateGraphicInSlides(ByVal GraphicName As String)

    Dim Presentation As Presentation: Set Presentation = App.ActivePresentationInstance
    
    Dim SourceShape As Shape
    If Not TryGetShapeByName(App.DesignSlide, GraphicName, SourceShape) Then Exit Sub
    
    Dim SlideIndex As Long
    For SlideIndex = App.ThemePreviewer.LayoutDesignsFirstSlideIndex To Presentation.Slides.count
        
        Dim CurrentSlide As Slide: Set CurrentSlide = Presentation.Slides(SlideIndex)
        Dim ExistingShape As Shape
        
        If TryGetShapeByName(CurrentSlide, GraphicName, ExistingShape) Then
            
            Dim CopyAttemptCount As Long: CopyAttemptCount = 0
            
            Do
                On Error Resume Next
                Err.Clear
                SourceShape.Copy
                Dim CopyFailed As Boolean: CopyFailed = (Err.Number <> 0)
                On Error GoTo 0
                
                If Not CopyFailed Then Exit Do
                
                Sleep 10
                DoEvents
                
                CopyAttemptCount = CopyAttemptCount + 1
                If CopyAttemptCount > 20 Then Exit Sub
            Loop
            
            Dim NewShape As Shape
            Dim PasteAttemptCount As Long: PasteAttemptCount = 0
            
            Do
                On Error Resume Next
                Err.Clear
                Set NewShape = CurrentSlide.Shapes.Paste(1)
                Dim PasteFailed As Boolean: PasteFailed = (Err.Number <> 0 Or NewShape Is Nothing)
                On Error GoTo 0
                
                If Not PasteFailed Then Exit Do
                
                Sleep 10
                DoEvents
                
                PasteAttemptCount = PasteAttemptCount + 1
                If PasteAttemptCount > 20 Then Exit Sub
            Loop
            
            If Not NewShape Is Nothing Then
                
                With NewShape
                    .LockAspectRatio = msoTrue
                    .Height = ExistingShape.Height
                    .Left = ExistingShape.Left
                    .Top = ExistingShape.Top
                    .Name = GraphicName
                End With
                
                ExistingShape.Delete
                
            End If
            
        End If
        
    Next SlideIndex

End Sub

Private Function TryGetShapeByName(ByVal Slide As Slide, ByVal ShapeName As String, ByRef FoundShape As Shape) As Boolean
    
    Dim ShapeIndex As Long
    For ShapeIndex = 1 To Slide.Shapes.count
        
        If Slide.Shapes(ShapeIndex).Name = ShapeName Then
            Set FoundShape = Slide.Shapes(ShapeIndex)
            TryGetShapeByName = True
            Exit Function
        End If
        
    Next ShapeIndex
    
End Function

Private Sub ApplySlideBackgroundToLayout(ByRef SourceSlide As PowerPoint.Slide, ByRef TargetLayout As PowerPoint.CustomLayout)
    TargetLayout.FollowMasterBackground = msoFalse
    Select Case SourceSlide.Background.Fill.Type
        Case msoFillSolid:
            ApplySolidBackgroundToLayout SourceSlide, TargetLayout
        Case msoFillPatterned:
            ApplyPatternFillBackgroundToLayout SourceSlide, TargetLayout
        Case msoFillBackground:
        Case msoFillPicture, msoFillTextured, msoFillGradient:
            If App.ShouldIgnoreComplexBackgrounds = False Then ApplyBackgroundThroughManualCopyFormat SourceSlide
    End Select
    If App.IsWaitingForLayoutClick = True Then Exit Sub
End Sub
Private Sub ApplyBackgroundThroughManualCopyFormat(ByRef SourceSlide As PowerPoint.Slide)
    
    Dim TargetLayout As PowerPoint.CustomLayout
    Set TargetLayout = SourceSlide.CustomLayout
    
    App.IsFirstBackgroundCopyExecution = False

    Dim ActiveDocumentWindow As DocumentWindow: Set ActiveDocumentWindow = ActiveWindow
    ActiveDocumentWindow.Selection.Unselect

    Dim ActivePresentationInstance As Presentation: Set ActivePresentationInstance = ActivePresentation

    ActiveDocumentWindow.View.GotoSlide SourceSlide.SlideIndex

    Dim ActivePane As Pane: Set ActivePane = ActiveDocumentWindow.Panes(1)
    ActivePane.Activate

    Application.CommandBars.ExecuteMso "FormatPainter"

    Dim WaitUntilTime As Double: WaitUntilTime = Timer + 0.3
    Do While Timer < WaitUntilTime
        DoEvents
    Loop

    ActiveDocumentWindow.ViewType = ppViewMasterThumbnails

    TargetLayout.Select

    App.IsWaitingForLayoutClick = True

    CopyBackground.Show

    If App.IsWaitingForLayoutClick Then Exit Sub

End Sub
Private Sub ApplySolidBackgroundToLayout(ByRef SourceSlide As PowerPoint.Slide, ByRef TargetLayout As PowerPoint.CustomLayout)
    TargetLayout.Background.Fill.Solid
    TargetLayout.FollowMasterBackground = msoFalse

    Dim srcFill As PowerPoint.FillFormat
    Dim dstFill As PowerPoint.FillFormat

    Set srcFill = SourceSlide.Background.Fill
    Set dstFill = TargetLayout.Background.Fill
    dstFill.Visible = srcFill.Visible
    dstFill.Transparency = srcFill.Transparency
    dstFill.ForeColor.Rgb = srcFill.ForeColor.Rgb
    dstFill.BackColor.Rgb = srcFill.BackColor.Rgb
End Sub

Private Sub CopyColorFormat(ByRef SourceColor As Variant, ByRef TargetColor As Variant)
    Select Case SourceColor.Type

        Case msoColorTypeRGB
            TargetColor.Rgb = SourceColor.Rgb

        Case msoColorTypeScheme
            TargetColor.ObjectThemeColor = SourceColor.ObjectThemeColor
            TargetColor.TintAndShade = SourceColor.TintAndShade
            TargetColor.Brightness = SourceColor.Brightness

        Case msoColorTypeCMYK, msoColorTypeCMS, msoColorTypeInk, msoColorTypeMixed
            TargetColor.Rgb = SourceColor.Rgb

        Case Else
            TargetColor.Rgb = SourceColor.Rgb

    End Select
End Sub
Private Sub ApplyPatternFillBackgroundToLayout(ByVal SourceSlide As PowerPoint.Slide, ByVal TargetLayout As PowerPoint.CustomLayout)

    TargetLayout.Background.Fill.Patterned SourceSlide.Background.Fill.Pattern

    CopyColorFormat SourceSlide.Background.Fill.ForeColor, TargetLayout.Background.Fill.ForeColor
    CopyColorFormat SourceSlide.Background.Fill.BackColor, TargetLayout.Background.Fill.BackColor

End Sub



Private Sub SyncLayoutFromSampleSlideShapes(ByRef SampleSlide As PowerPoint.Slide)
    On Error Resume Next
    'SampleSlide.CustomLayout.Select
    On Error GoTo 0
    If SampleSlide Is Nothing Then Exit Sub
    If SampleSlide.CustomLayout Is Nothing Then Exit Sub

    SyncPlaceholderShapes SampleSlide
    CopyNonPlaceholderShapesToLayout SampleSlide

End Sub
Private Sub SyncPlaceholderShapes(ByRef SampleSlide As PowerPoint.Slide)

    Dim TargetLayout As PowerPoint.CustomLayout: Set TargetLayout = SampleSlide.CustomLayout
    

    Dim SourceShape As PowerPoint.Shape
    For Each SourceShape In SampleSlide.Shapes

        If Not IsPlaceholderShape(SourceShape) Then GoTo ContinueLoop

        Dim DestinationShape As PowerPoint.Shape
        Set DestinationShape = GetShapeByName(TargetLayout.Shapes, SourceShape.Name)

        If Not DestinationShape Is Nothing Then
            ApplyGeometryFromShape SourceShape, DestinationShape
            ApplyFormattingFromShape SourceShape, DestinationShape
            ApplyTextFormattingFromShape SourceShape, DestinationShape
        End If

ContinueLoop:
    Next SourceShape

End Sub
Private Sub CopyNonPlaceholderShapesToLayout(ByRef SampleSlide As PowerPoint.Slide)

    Dim NonPlaceholderShapeCountInSlide As Integer: NonPlaceholderShapeCountInSlide = GetNonPlaceholderShapeCount(SampleSlide.Shapes)

RetryCopyProcess:

    Dim TargetLayout As PowerPoint.CustomLayout: Set TargetLayout = SampleSlide.CustomLayout

    Dim ShapesToCopy As PowerPoint.ShapeRange: Set ShapesToCopy = GetNonPlaceholderShapeRange(SampleSlide.Shapes)

    If ShapesToCopy Is Nothing Then Exit Sub

    Dim ShapeCount As Long: ShapeCount = ShapesToCopy.count

    PasteShapeRangeToLayout ShapesToCopy, TargetLayout

    If ShouldRetryCopyProcess(TargetLayout, ShapeCount) Then GoTo RetryCopyProcess

End Sub
Private Function ShouldRetryCopyProcess(ByRef TargetLayout As PowerPoint.CustomLayout, ByVal ExpectedShapeCount As Long) As Boolean

    Dim LayoutNonPlaceholderCount As Long: LayoutNonPlaceholderCount = GetNonPlaceholderShapeCount(TargetLayout.Shapes)

    If LayoutNonPlaceholderCount = ExpectedShapeCount Then Exit Function

    Dim ShapeIndex As Long
    For ShapeIndex = TargetLayout.Shapes.count To 1 Step -1
        If Not IsPlaceholderShape(TargetLayout.Shapes(ShapeIndex)) Then
            TargetLayout.Shapes(ShapeIndex).Delete
        End If
    Next ShapeIndex

    Sleep 20
    DoEvents

    ShouldRetryCopyProcess = True

End Function
Private Function GetNonPlaceholderShapeRange(ByRef Shapes As PowerPoint.Shapes) As PowerPoint.ShapeRange

    Dim ShapeIndexes() As Long
    Dim ShapeCount As Long: ShapeCount = 0

    Dim ShapeIndex As Long
    For ShapeIndex = 1 To Shapes.count
        If Not IsPlaceholderShape(Shapes(ShapeIndex)) Then
            ShapeCount = ShapeCount + 1
            ReDim Preserve ShapeIndexes(1 To ShapeCount)
            ShapeIndexes(ShapeCount) = ShapeIndex
        End If
    Next ShapeIndex

    If ShapeCount = 0 Then Exit Function

    Set GetNonPlaceholderShapeRange = Shapes.Range(ShapeIndexes)
    
'    '/* DEBUG
'    Dim ResultShapeRange As PowerPoint.ShapeRange
'    Set ResultShapeRange = GetNonPlaceholderShapeRange
'    Debug.Print "Range shapes to copy"
'    Dim ShapeIndexInRange As Long
'    For ShapeIndexInRange = 1 To ResultShapeRange.count
'        Debug.Print ResultShapeRange(ShapeIndexInRange).Name
'    Next ShapeIndexInRange

End Function

Private Sub PasteShapeRangeToLayout(ByRef ShapesToCopy As PowerPoint.ShapeRange, ByRef TargetLayout As PowerPoint.CustomLayout)

CopyShapeSafely:

    On Error GoTo CopyShapeNotReady
    ShapesToCopy.Copy
    DoEvents
    On Error GoTo 0

PasteShapeSafely:

    On Error GoTo PasteShapeNotReady
    Dim PastedShapes As PowerPoint.ShapeRange: Set PastedShapes = TargetLayout.Shapes.Paste
    DoEvents
    On Error GoTo 0

    Exit Sub

CopyShapeNotReady:
    Sleep 10
    Resume CopyShapeSafely

PasteShapeNotReady:
    Sleep 10
    Resume PasteShapeSafely

End Sub
Private Function GetNonPlaceholderShapeCount(ByRef Shapes As PowerPoint.Shapes) As Long

    Dim ShapeIndex As Long
    For ShapeIndex = 1 To Shapes.count
        If Not IsPlaceholderShape(Shapes(ShapeIndex)) Then
            GetNonPlaceholderShapeCount = GetNonPlaceholderShapeCount + 1
        End If
    Next ShapeIndex

End Function

Public Function GetShapeByName(ByRef Shapes As PowerPoint.Shapes, ByVal ShapeName As String) As PowerPoint.Shape
    If Len(ShapeName) = 0 Then Exit Function

    On Error Resume Next
    Set GetShapeByName = Shapes(ShapeName)
    On Error GoTo 0
End Function

Private Sub ApplyGeometryFromShape(ByRef Source As PowerPoint.Shape, ByRef Destination As PowerPoint.Shape)
    Destination.Left = Source.Left
    Destination.Top = Source.Top
    Destination.Width = Source.Width
    Destination.Height = Source.Height
    Destination.Rotation = Source.Rotation
End Sub

Private Sub ApplyFormattingFromShape(ByRef Source As PowerPoint.Shape, ByRef Destination As PowerPoint.Shape)
    On Error Resume Next

    Destination.Fill.Visible = Source.Fill.Visible
    Destination.Fill.ForeColor.Rgb = Source.Fill.ForeColor.Rgb
    Destination.Fill.Transparency = Source.Fill.Transparency
    Destination.Fill.Solid

    Destination.Line.Visible = Source.Line.Visible
    Destination.Line.ForeColor.Rgb = Source.Line.ForeColor.Rgb
    Destination.Line.Transparency = Source.Line.Transparency
    Destination.Line.Weight = Source.Line.Weight
    Destination.Line.DashStyle = Source.Line.DashStyle

    Destination.Shadow.Visible = Source.Shadow.Visible
    If Destination.Shadow.Visible = msoTrue Then
        Destination.Shadow.ForeColor.Rgb = Source.Shadow.ForeColor.Rgb
        Destination.Shadow.Transparency = Source.Shadow.Transparency
        Destination.Shadow.Blur = Source.Shadow.Blur
        Destination.Shadow.OffsetX = Source.Shadow.OffsetX
        Destination.Shadow.OffsetY = Source.Shadow.OffsetY
    End If

    On Error GoTo 0
End Sub

Private Sub ApplyTextFormattingFromShape(ByRef Source As PowerPoint.Shape, ByRef Destination As PowerPoint.Shape)
    If Not Source.HasTextFrame Then Exit Sub
    If Not Destination.HasTextFrame Then Exit Sub

    If Not Source.TextFrame.HasText Then Exit Sub

    On Error Resume Next

    Destination.TextFrame.MarginLeft = Source.TextFrame.MarginLeft
    Destination.TextFrame.MarginRight = Source.TextFrame.MarginRight
    Destination.TextFrame.MarginTop = Source.TextFrame.MarginTop
    Destination.TextFrame.MarginBottom = Source.TextFrame.MarginBottom

    Destination.TextFrame.VerticalAnchor = Source.TextFrame.VerticalAnchor

    Destination.TextFrame.TextRange.ParagraphFormat.Alignment = Source.TextFrame.TextRange.ParagraphFormat.Alignment
    Destination.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = Source.TextFrame.TextRange.ParagraphFormat.Bullet.Visible

    Destination.TextFrame.TextRange.Font.Name = Source.TextFrame.TextRange.Font.Name
    Destination.TextFrame.TextRange.Font.size = Source.TextFrame.TextRange.Font.size
    Destination.TextFrame.TextRange.Font.Bold = Source.TextFrame.TextRange.Font.Bold
    Destination.TextFrame.TextRange.Font.Italic = Source.TextFrame.TextRange.Font.Italic
    Destination.TextFrame.TextRange.Font.Underline = Source.TextFrame.TextRange.Font.Underline
    Destination.TextFrame.TextRange.Font.Color.Rgb = Source.TextFrame.TextRange.Font.Color.Rgb

    On Error GoTo 0
End Sub

Public Sub ClearNonPlaceholderShapesFromLayoutsInLayoutDesignSection()

    If App Is Nothing Then Set App = New App

    Dim presentationInstance As PowerPoint.Presentation
    Set presentationInstance = App.ActivePresentationInstance

    Dim SectionIndex As Long
    SectionIndex = GetSectionIndexByName(presentationInstance, "Layout Designs")
    If SectionIndex = -1 Then Exit Sub

    Dim FirstSlideIndex As Long
    FirstSlideIndex = presentationInstance.SectionProperties.FirstSlide(SectionIndex)

    Dim SlideCount As Long
    SlideCount = presentationInstance.SectionProperties.SlidesCount(SectionIndex)

    Dim SlideIndex As Long

    Dim ProcessedLayouts As Object
    Set ProcessedLayouts = CreateObject("Scripting.Dictionary")

    For SlideIndex = FirstSlideIndex To FirstSlideIndex + SlideCount - 1
        
        Dim CurrentSlide As PowerPoint.Slide
        Set CurrentSlide = presentationInstance.Slides(SlideIndex)

        Dim LayoutId As Long: LayoutId = CurrentSlide.CustomLayout.Index

        If Not ProcessedLayouts.Exists(LayoutId) Then

            On Error Resume Next
            CurrentSlide.CustomLayout.Select
            On Error GoTo 0

            ClearNonPlaceholderShapesFromLayout CurrentSlide.CustomLayout
            ProcessedLayouts.Add LayoutId, CurrentSlide.CustomLayout

        End If

    Next SlideIndex
    
    VerifyLayoutsContainOnlyPlaceholders ProcessedLayouts

End Sub

Public Sub ClearNonPlaceholderShapesFromLayout(ByRef TargetLayout As PowerPoint.CustomLayout)

    Dim i As Long

    For i = TargetLayout.Shapes.count To 1 Step -1

        Dim shp As PowerPoint.Shape
        Set shp = TargetLayout.Shapes(i)

        If Not IsPlaceholderShape(shp) Then
            'shp.Select '/**
            'Debug.Print shp.Name
            shp.Delete
 
        End If

    Next i
End Sub
Public Function GetSectionIndexByName(ByRef presentationInstance As PowerPoint.Presentation, ByRef SectionName As String) As Long

    Dim i As Long

    For i = 1 To presentationInstance.SectionProperties.count
        If presentationInstance.SectionProperties.Name(i) = SectionName Then
            GetSectionIndexByName = i
            Exit Function
        End If
    Next i

    GetSectionIndexByName = -1

End Function
Public Function IsPlaceholderShape(ByRef Shape As PowerPoint.Shape) As Boolean

    If Shape Is Nothing Then Exit Function

    Err.Clear
    On Error Resume Next
    'Shape.Select 'solo para debugueo
    On Error GoTo 0

    If Shape.Type = msoPlaceholder Then IsPlaceholderShape = True: Exit Function

    If InStr(1, Shape.Name, "Placeholder", vbTextCompare) > 0 Then
        IsPlaceholderShape = True
        Exit Function
    End If

    Err.Clear
    On Error Resume Next
    Dim PlaceholderType As Long: PlaceholderType = CLng(Shape.PlaceholderFormat.Type)
    If Err.Number = 0 Then IsPlaceholderShape = True
    On Error GoTo 0

End Function

Public Function IsBrandlogoShape(ByRef Shape As PowerPoint.Shape) As Boolean
    If Shape Is Nothing Then Exit Function
    
    Dim ShapeName As String
    ShapeName = Trim$(Shape.Name)
    
    Select Case ShapeName
        Case "BrandLogoImage", "PositiveMark", "NegativeMark", "NegativeLogo"
            IsBrandlogoShape = True
        Case Else
            IsBrandlogoShape = False
    End Select
End Function


Private Sub VerifyLayoutsContainOnlyPlaceholders(ByRef ProcessedLayouts As Object)

    Dim LayoutKey As Variant

    For Each LayoutKey In ProcessedLayouts.Keys

        Dim TargetLayout As PowerPoint.CustomLayout
        Set TargetLayout = ProcessedLayouts(LayoutKey)

        Dim ShapeIndex As Long

        For ShapeIndex = 1 To TargetLayout.Shapes.count
            If Not IsPlaceholderShape(TargetLayout.Shapes(ShapeIndex)) Then
                Debug.Print "Non-placeholder found in layout: " & TargetLayout.Name
                Exit Sub
            End If
        Next ShapeIndex

    Next LayoutKey

End Sub

Public Function AllNonPlaceholderShapesExistInLayout(ByRef SampleSlide As PowerPoint.Slide) As Boolean

    If SampleSlide Is Nothing Then Exit Function
    If SampleSlide.CustomLayout Is Nothing Then Exit Function

    Dim TargetLayout As PowerPoint.CustomLayout
    Set TargetLayout = SampleSlide.CustomLayout

    Dim SourceShape As PowerPoint.Shape

    For Each SourceShape In SampleSlide.Shapes

        If IsPlaceholderShape(SourceShape) Then GoTo ContinueLoop

        Dim LayoutShape As PowerPoint.Shape
        Set LayoutShape = GetShapeByName(TargetLayout.Shapes, SourceShape.Name)

        If LayoutShape Is Nothing Then Exit Function

ContinueLoop:
    Next SourceShape

    AllNonPlaceholderShapesExistInLayout = True

End Function



Private Property Get BrandLogoImage() As PowerPoint.Shape
    Set BrandLogoImage = App.Slide.Shapes("BrandLogoImage")
End Property
