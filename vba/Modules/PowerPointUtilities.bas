Attribute VB_Name = "PowerPointUtilities"

#If VBA7 Then
    Private Declare PtrSafe Function LCIDToLocaleName Lib "kernel32" (ByVal Locale As Long, ByVal lpName As LongPtr, ByVal cchName As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function GetLocaleInfoEx Lib "kernel32" (ByVal lpLocaleName As LongPtr, ByVal LCType As Long, ByVal lpLCData As LongPtr, ByVal cchData As Long) As Long
#Else
    Private Declare Function LCIDToLocaleName Lib "kernel32" (ByVal Locale As Long, ByVal lpName As Long, ByVal cchName As Long, ByVal dwFlags As Long) As Long
    Private Declare Function GetLocaleInfoEx Lib "kernel32" (ByVal lpLocaleName As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
#End If
Private Const LOCALE_SLOCALIZEDDISPLAYNAME As Long = &H2

Option Explicit
Option Private Module

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Public Sub ShowBackGroundFillColorPicker()
    PowerPoint.Application.CommandBars.ExecuteMso "EyedropperFill"
End Sub

Public Sub ShowTextFillColorPicker()
    PowerPoint.Application.CommandBars.ExecuteMso "EyedropperFillText"
End Sub

Public Sub ShowAltText()
    If PowerPoint.Application.CommandBars.GetPressedMso("AltTextPaneRibbon") = False Then
        PowerPoint.Application.CommandBars.ExecuteMso "AltTextPaneRibbon"
    End If
End Sub

Public Function AreTheSameShape(ByRef Selection As PowerPoint.Selection, ByRef Shape As PowerPoint.Shape) As Boolean
    Dim Result As Boolean
    
    Result = True
    Result = (Selection.Type = ppSelectionShapes Or Selection.Type = ppSelectionText):      If Result = False Then GoSub CleanExit
    Result = Result And Selection.ShapeRange.count = 1:                                     If Result = False Then GoSub CleanExit
    Result = Result And (Selection.ShapeRange(1).Id = Shape.Id):                            If Result = False Then GoSub CleanExit
    Result = Result And (Selection.SlideRange(1).SlideID = Shape.Parent.SlideID):           If Result = False Then GoSub CleanExit
    
CleanExit:
    AreTheSameShape = Result
End Function

Public Function GetSlideByTitleIn(ByRef Presentation As PowerPoint.Presentation, ByVal SlideName As String) As PowerPoint.Slide
    Dim Result As PowerPoint.Slide
    For Each Result In Presentation.Slides
        If Result.Shapes.Title.TextFrame2.TextRange.Text = SlideName Then Exit For
    Next Result
    
    Set GetSlideByTitleIn = Result
End Function

Public Function GetFirstParagraphOf(ByRef Range As PowerPoint.TextRange) As TextRange
    Dim TextFrame As PowerPoint.TextFrame
    Set TextFrame = Range.Parent
    
    Dim n As Long
    For n = 1 To TextFrame.TextRange.Paragraphs.count
        If TextFrame.TextRange.Paragraphs(n, 1).start > Range.start Then Exit For
    Next n
    
    Set GetFirstParagraphOf = TextFrame.TextRange.Paragraphs(n - 1, 1)
End Function

Public Sub ApplyFontToShape(ByRef Shape As PowerPoint.Shape, ByRef Font As Variant)
    With Shape.TextFrame.TextRange.Font
        Select Case True
            Case VarType(Font) = vbString:                  .Name = Font
            Case TypeOf Font Is PowerPoint.Font:            .Name = Font.Name
        End Select
    End With
End Sub

'Public Sub SetMastersDefaultShape(ByRef Source As PowerPoint.Shape, ByRef Destination As PowerPoint.Master)
'
'CopyShapeSafely:
'
'    On Error GoTo CopyShapeNotReady
'    Source.Copy
'    DoEvents
'    On Error GoTo 0
'
'PasteShapeSafely:
'    On Error GoTo PasteShapeNotReady
'    Dim DefaultShape As PowerPoint.Shape
'    Set DefaultShape = Destination.Shapes.Paste(1) 'It should happen in Master Slides
'    DoEvents
'    On Error GoTo 0
'
'    DefaultShape.SetShapesDefaultProperties
'    DefaultShape.Delete
'
'Exit Sub
'
'CopyShapeNotReady:
'    Sleep 10
'    Resume CopyShapeSafely:
'
'
'PasteShapeNotReady:
'    Sleep 10
'    Resume PasteShapeSafely:
'
'End Sub

Public Sub SetMastersDefaultShape(ByRef Source As PowerPoint.Shape, ByRef Destination As PowerPoint.Master)

    Dim CopyAttemptCount As Long: CopyAttemptCount = 0
    
    Do
        On Error Resume Next
        Source.Copy
        On Error GoTo 0
        
        If Err.Number = 0 Then Exit Do
        
        Sleep 10
        DoEvents
        CopyAttemptCount = CopyAttemptCount + 1
        
        If CopyAttemptCount > 20 Then Exit Sub
    Loop
    
    
    Dim DefaultShape As PowerPoint.Shape
    Dim PasteAttemptCount As Long: PasteAttemptCount = 0
    
    Do
        On Error Resume Next
        Set DefaultShape = Destination.Shapes.Paste(1)
        On Error GoTo 0
        
        If Not DefaultShape Is Nothing Then Exit Do
        
        Sleep 10
        DoEvents
        PasteAttemptCount = PasteAttemptCount + 1
        
        If PasteAttemptCount > 20 Then Exit Sub
    Loop
    
    DefaultShape.SetShapesDefaultProperties
    
    DefaultShape.Delete

End Sub

Public Function GetLastCharacterInNameOf(ByRef Shape As PowerPoint.Shape) As String
    GetLastCharacterInNameOf = Right(Shape.Name, 1)
End Function

Public Sub CleanUnusedThemeDesignsInActivePresentation()
    On Error GoTo CleanFail

    Dim Presentation As PowerPoint.Presentation
    Set Presentation = ActivePresentation

    If Presentation Is Nothing Then
        MsgBox "No active presentation found.", vbExclamation
        Exit Sub
    End If

    Dim InitialDesignCount As Long
    InitialDesignCount = Presentation.Designs.count

    Dim DeletedDesigns As Long
    Dim LockedDesigns As Long

    Dim DesignIndex As Long
    For DesignIndex = Presentation.Designs.count To 2 Step -1
        DeleteDesignIfPossible Presentation, DesignIndex, DeletedDesigns, LockedDesigns
    Next DesignIndex

    MsgBox "Theme cleanup completed." & vbCrLf & _
           "Initial designs: " & InitialDesignCount & vbCrLf & _
           "Deleted designs: " & DeletedDesigns & vbCrLf & _
           "Designs still in use or protected: " & LockedDesigns, vbInformation

    Exit Sub

CleanFail:
    MsgBox "Unexpected error while cleaning designs: " & Err.Description, vbCritical
End Sub

Private Sub DeleteDesignIfPossible(ByRef Presentation As PowerPoint.Presentation, _
                                   ByVal DesignIndex As Long, _
                                   ByRef DeletedDesigns As Long, _
                                   ByRef LockedDesigns As Long)
    On Error GoTo DeleteFailed

    Presentation.Designs(DesignIndex).Delete
    DeletedDesigns = DeletedDesigns + 1

    Exit Sub

DeleteFailed:
    LockedDesigns = LockedDesigns + 1
    Err.Clear
End Sub


Public Function GetFullLanguageNameFromLanguageId(ByVal LanguageId As MsoLanguageID) As String

    Dim LocaleBuffer As String: LocaleBuffer = String$(85, vbNullChar)

    Dim CharactersWritten As Long
    CharactersWritten = LCIDToLocaleName(LanguageId, StrPtr(LocaleBuffer), Len(LocaleBuffer), 0)

    If CharactersWritten = 0 Then
        GetFullLanguageNameFromLanguageId = CStr(LanguageId)
        Exit Function
    End If

    Dim LocaleName As String: LocaleName = Left$(LocaleBuffer, CharactersWritten - 1)

    Dim NameBuffer As String: NameBuffer = String$(85, vbNullChar)

    Dim ResultLength As Long
    ResultLength = GetLocaleInfoEx(StrPtr(LocaleName), LOCALE_SLOCALIZEDDISPLAYNAME, StrPtr(NameBuffer), Len(NameBuffer))

    If ResultLength > 0 Then
        GetFullLanguageNameFromLanguageId = Left$(NameBuffer, ResultLength - 1)
    Else
        GetFullLanguageNameFromLanguageId = LocaleName
    End If

End Function
