VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppThemeCreator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private BrandThemeName As String
Private BrandThemePath As String

Private UnreliableThemeName As String
Private UnreliableThemePath As String

Friend Sub CreateTheme(ByVal OpenFolder As Boolean, Optional ByVal TimeStamp As String = "")
    Dim BrandName As String
    BrandName = App.Slide.Shapes("BrandName").TextFrame.TextRange.Text

    If Len(TimeStamp) = 0 Then
        TimeStamp = Format(Now, "YYYY.MM.DD.hhnn")
    End If

    Dim DestinationFolderPath As String
    DestinationFolderPath = GetKnownFolderPath(kfID_Downloads) & "\"
    
    Dim folderName As String
    folderName = BrandName & "'s Office files - " & TimeStamp
    
    Dim folderPath As String
    folderPath = CreateBrandFolder(DestinationFolderPath, folderName)
    
    FontsToThemeBack ActivePresentation.Designs(1).SlideMaster
            
    If App.BothBrandFontsAreInWindowsAndMacFontsList Then
        
        Dim fileName As String
        fileName = BrandName & "'s Office theme - " & TimeStamp & ".thmx"
        BrandThemePath = ExportCurrentPresentationAsTheme(BrandName, folderPath, TimeStamp, fileName)
        UnreliableThemePath = ""
    
    Else

        
        Dim ReliableFileName As String
        ReliableFileName = BrandName & "'s Office theme - Reliable Fonts - " & TimeStamp & ".thmx"
        BrandThemePath = ExportCurrentPresentationAsTheme(BrandName, folderPath, TimeStamp, ReliableFileName)
        
        ApplyCorporateFontsToTheme ActivePresentation.Designs(1).SlideMaster
        
        Dim UnreliableFileName As String
        UnreliableFileName = BrandName & "'s Office theme - Unreliable Fonts - " & TimeStamp & ".thmx"
        UnreliableThemePath = ExportCurrentPresentationAsTheme(BrandName, folderPath, TimeStamp, UnreliableFileName)
        
        FontsToThemeBack ActivePresentation.Designs(1).SlideMaster
        
    End If
    If OpenFolder Then
        'If MsgBox("Open theme?", vbYesNo) = vbYes Then
            VBA.Shell "explorer.exe /select, " & BrandThemePath, vbMaximizedFocus
        'End If
    End If

End Sub

Friend Function ExportCurrentPresentationAsTheme( _
    ByVal BrandName As String, _
    ByVal DestinationFolder As String, _
    ByVal TimeStamp, _
    ByVal fileName As String) As String

    Dim fullPath As String
    fullPath = DestinationFolder & "\" & fileName

    App.Presentation.SaveCopyAs fullPath, ppSaveAsOpenXMLTheme
    ExportCurrentPresentationAsTheme = fullPath
End Function

Friend Function ExportThemeWithCorporateFonts( _
    ByVal BrandName As String, _
    ByVal folderPath As String, AppendTimestamp) As String

    Dim ThemeFileName As String
    ThemeFileName = BrandName & "'s Office theme - Unreliable Fonts - " & AppendTimestamp & ".thmx"
    
    ApplyCorporateFontsToTheme ActivePresentation.Designs(1).SlideMaster
    Dim ThemePath As String
    ThemePath = folderPath & "\" & ThemeFileName
    App.Presentation.SaveCopyAs ThemePath, ppSaveAsOpenXMLTheme
    FontsToThemeBack ActivePresentation.Designs(1).SlideMaster
    
    ExportThemeWithCorporateFonts = ThemePath
End Function

Private Function CreateBrandFolder(ByVal ParentFolderPath As String, ByVal BrandFileName As String) As String
    Dim FullFolderPath As String
    FullFolderPath = ParentFolderPath & BrandFileName

    If Dir(FullFolderPath, vbDirectory) = vbNullString Then MkDir FullFolderPath

    CreateBrandFolder = FullFolderPath
End Function

Private Sub ApplyCorporateFontsToTheme(ByRef Master As PowerPoint.Master)
    With Master.Theme.ThemeFontScheme
        .MajorFont.Item(msoThemeLatin).Name = App.Slide.Shapes("HeadingsBrandFontNameTextbox").TextFrame.TextRange.Text
        .MinorFont.Item(msoThemeLatin).Name = App.Slide.Shapes("BodyBrandFontNameTextbox").TextFrame.TextRange.Text
    End With
End Sub
Private Sub FontsToThemeBack(ByRef Master As PowerPoint.Master)
    With Master.Theme.ThemeFontScheme
        .MajorFont.Item(msoThemeLatin).Name = App.HeadingsAlternativeFontNameTextbox.TextFrame.TextRange.Text
        
        .MinorFont.Item(msoThemeLatin).Name = App.BodyAlternativeFontNameTextbox.TextFrame.TextRange.Text
    End With
End Sub

'Friend Property Get GetBrandThemeName() As String
'    GetBrandThemeName = BrandThemeName
'End Property

Friend Property Get GetBrandThemePath() As String
    GetBrandThemePath = BrandThemePath
End Property

'Friend Property Get GetUnreliableThemeName() As String
'    GetUnreliableThemeName = UnreliableThemeName
'End Property

Friend Property Get GetUnreliableThemePath() As String
    GetUnreliableThemePath = UnreliableThemePath
End Property

Private Function GetFileName(ByVal FilePath As String) As String
    GetFileName = CreateObject("Scripting.FileSystemObject").GetFileName(FilePath)
End Function



