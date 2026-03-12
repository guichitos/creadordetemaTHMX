VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppThemeApplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Friend Sub ApplyTheme()
    
    Dim TimeStamp As String
    TimeStamp = Format(Now, "YYYY.MM.DD.hhnn")
    
    App.ThemeCreator.CreateTheme False, TimeStamp
    
    Dim BrandName As String
    BrandName = App.Slide.Shapes("BrandName").TextFrame.TextRange.Text

    Dim DestinationFolderPath As String
    DestinationFolderPath = GetKnownFolderPath(kfID_Downloads) & "\"

    Dim folderName As String
    folderName = BrandName & "'s Office files - " & TimeStamp
    DataEntryForm.Show
End Sub

Public Sub CreateTemplatesFromThemeIn(ByVal ThemeFilePath As String, ByVal DestinationFolderPath As String)
    CreateExcelTemplate ThemeFilePath, DestinationFolderPath
    CreateWordTemplate ThemeFilePath, DestinationFolderPath
    CreatePowerPointTemplate ThemeFilePath, DestinationFolderPath
End Sub

Private Sub CreateExcelTemplate(ByVal ThemeFilePath As String, ByVal DestinationFolderPath As String)
    Dim ExcelApplication As Object
    Set ExcelApplication = CreateObject("Excel.Application")
    ExcelApplication.Visible = App.IsInDesign

    Dim ExcelWorkbook As Object
    Set ExcelWorkbook = ExcelApplication.Workbooks.Add

    ExcelWorkbook.ApplyTheme ThemeFilePath

    Dim ExcelFilePath As String
    ExcelFilePath = DestinationFolderPath & "\" & "Spreadsheet template" & ".xltx"
    
    ExcelWorkbook.SaveAs ExcelFilePath, 54

    ExcelWorkbook.Close False
    ExcelApplication.Quit

    Set ExcelWorkbook = Nothing
    Set ExcelApplication = Nothing
End Sub

Private Sub CreateWordTemplate(ByVal ThemeFilePath As String, ByVal DestinationFolderPath As String)
    Dim WordApplication As Word.Application
    Set WordApplication = CreateObject("Word.Application")
    WordApplication.Visible = App.IsInDesign

    Dim WordDocument As Object
    Set WordDocument = WordApplication.Documents.Add
    
    WordDocument.ApplyDocumentTheme ThemeFilePath

    Dim WordFilePath As String
    WordFilePath = DestinationFolderPath & "\" & "Document template" & ".dotx"
    WordDocument.SaveAs2 WordFilePath, 14

    WordDocument.Close False
    WordApplication.Quit

    Set WordDocument = Nothing
    Set WordApplication = Nothing
End Sub

Private Sub CreatePowerPointTemplate(ByVal ThemeFilePath As String, ByVal DestinationFolderPath As String)
    Dim PowerPointApplication As Object
    Set PowerPointApplication = CreateObject("PowerPoint.Application")
    
    PowerPointApplication.WindowState = ppWindowMinimized

    Dim PowerPointPresentation As Object
    Set PowerPointPresentation = PowerPointApplication.Presentations.Add

    PowerPointPresentation.ApplyTheme ThemeFilePath

    Dim PowerPointFilePath As String
    PowerPointFilePath = DestinationFolderPath & "\" & "Presentation template" & ".potx"
    PowerPointPresentation.SaveAs PowerPointFilePath, ppSaveAsOpenXMLTemplate

    PowerPointPresentation.Close
'    On Error Resume Next
'    PowerPointApplication.Quit
'    On Error GoTo 0

    Set PowerPointPresentation = Nothing
'    Set PowerPointApplication = Nothing
End Sub

Public Sub ApplyThemeToFolderFiles( _
    ByRef WordApp As Word.Application, _
    ByRef ExcelApp As Excel.Application, _
    ByRef PowerPointApp As PowerPoint.Application, _
    ByVal ThemePath As String, _
    ByVal SourceFolder As Object, _
    ByVal TargetFolderPath As String, _
    ByVal FileSystem As Object)

    Dim FileItem As Object
    For Each FileItem In SourceFolder.Files
        Dim fileName As String: fileName = FileItem.Name
        Debug.Print fileName
        If Left(fileName, 2) <> "~$" Then
            Dim FileExtension As String: FileExtension = LCase(FileSystem.GetExtensionName(fileName))
            Dim OutputPath As String: OutputPath = TargetFolderPath & "\" & fileName

            Select Case FileExtension
                Case "docx", "docm", "dotx", "dotm"
                    App.ThemeApplier.ApplyThemeToExistingWordFile WordApp, ThemePath, FileItem.Path, OutputPath, FileExtension
                Case "xlsx", "xlsm", "xltx", "xltm"
                    App.ThemeApplier.ApplyThemeToExistingExcelFile ExcelApp, ThemePath, FileItem.Path, OutputPath, FileExtension
                Case "pptx", "pptm", "potx", "potm"
                    App.ThemeApplier.ApplyThemeToExistingPowerPointFile PowerPointApp, ThemePath, FileItem.Path, OutputPath, FileExtension
            End Select
        End If
    Next
End Sub

Public Function GetExcelFileFormat(ByVal Extension As String) As XlFileFormat
    Select Case LCase(Extension)
        Case "xlsx": GetExcelFileFormat = xlOpenXMLWorkbook
        Case "xlsm": GetExcelFileFormat = xlOpenXMLWorkbookMacroEnabled
        Case "xls":  GetExcelFileFormat = xlExcel8
        Case "xltx": GetExcelFileFormat = xlOpenXMLTemplate
        Case "xltm": GetExcelFileFormat = xlOpenXMLTemplateMacroEnabled
        Case Else
            Err.Raise vbObjectError + 1, "GetExcelFileFormat", _
                "Extensión de archivo no soportada: " & Extension
    End Select
End Function

Public Function GetWordFileFormat(ByVal Extension As String) As WdSaveFormat
    Select Case LCase(Extension)
        Case "docx": GetWordFileFormat = wdFormatXMLDocument
        Case "docm": GetWordFileFormat = wdFormatXMLDocumentMacroEnabled
        Case "doc":  GetWordFileFormat = wdFormatDocument
        Case "dotx": GetWordFileFormat = wdFormatXMLTemplate
        Case "dotm": GetWordFileFormat = wdFormatXMLTemplateMacroEnabled
        Case "dot":  GetWordFileFormat = wdFormatTemplate
        Case Else
            Err.Raise vbObjectError + 1, "GetWordFileFormat", _
                "Unsupported Word file extension: " & Extension
    End Select
End Function

Public Function GetPowerPointFileFormat(ByVal Extension As String) As PpSaveAsFileType
    Select Case LCase(Extension)
        Case "pptx": GetPowerPointFileFormat = ppSaveAsOpenXMLPresentation
        Case "pptm": GetPowerPointFileFormat = ppSaveAsOpenXMLPresentationMacroEnabled
        Case "potx": GetPowerPointFileFormat = ppSaveAsOpenXMLTemplate
        Case "potm": GetPowerPointFileFormat = ppSaveAsOpenXMLTemplateMacroEnabled
        Case "ppt":  GetPowerPointFileFormat = ppSaveAsPresentation
        Case "pot":  GetPowerPointFileFormat = ppSaveAsTemplate
        Case Else
            Err.Raise vbObjectError + 1, "GetPowerPointFileFormat", _
                "Unsupported PowerPoint file extension: " & Extension
    End Select
End Function

Public Sub ApplyThemeToExistingWordFile(ByVal WordApp As Word.Application, _
                                          ByVal ThemeFilePath As String, _
                                          ByVal SourceFilePath As String, _
                                          ByVal SaveAsPath As String, _
                                          ByVal ext As String)
    If ThemeFilePath <> "" Then
        Dim WordDocument As Word.Document
        Set WordDocument = WordApp.Documents.Open(SourceFilePath)
        WordDocument.ApplyDocumentTheme ThemeFilePath
        WordDocument.SaveAs2 fileName:=SaveAsPath, FileFormat:=GetWordFileFormat(ext)
        WordDocument.Close
        Set WordDocument = Nothing
    End If
End Sub

Public Sub ApplyThemeToExistingExcelFile(ByVal ExcelApp As Excel.Application, _
                                          ByVal ThemeFilePath As String, _
                                          ByVal SourceFilePath As String, _
                                          ByVal SaveAsPath As String, _
                                          ByVal ext As String)
    If ThemeFilePath <> "" Then
        Dim ExcelWorkbook As Excel.Workbook
        Set ExcelWorkbook = ExcelApp.Workbooks.Open(SourceFilePath)
        ExcelWorkbook.ApplyTheme ThemeFilePath
        ExcelWorkbook.SaveAs fileName:=SaveAsPath, FileFormat:=GetExcelFileFormat(ext)
        ExcelWorkbook.Close False
        Set ExcelWorkbook = Nothing
    End If
End Sub

Public Sub ApplyThemeToExistingPowerPointFile(ByVal PowerPointApp As PowerPoint.Application, _
                                               ByVal ThemeFilePath As String, _
                                               ByVal SourceFilePath As String, _
                                               ByVal SaveAsPath As String, _
                                               ByVal ext As String)
    Dim PowerPointPres As PowerPoint.Presentation
    Set PowerPointPres = PowerPointApp.Presentations.Open(SourceFilePath, WithWindow:=msoFalse)
    PowerPointPres.ApplyTheme ThemeFilePath
    PowerPointPres.SaveAs fileName:=SaveAsPath, FileFormat:=GetPowerPointFileFormat(ext)
    PowerPointPres.Close
    Set PowerPointPres = Nothing
End Sub

Public Sub ApplyThemesToFolder( _
    ByVal SourceFolderPath As String, _
    ByVal ThemePaths As Object)

    If Not IsValidFolderAndThemeSelection(SourceFolderPath, ThemePaths) Then Exit Sub

    Dim FileSystem As Object: Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Dim SourceFolder As Object: Set SourceFolder = FileSystem.GetFolder(SourceFolderPath)

    Dim CurrentFolderName As String
    CurrentFolderName = SourceFolder.Name

    Dim WordApp As Word.Application: Set WordApp = New Word.Application
    Dim ExcelApp As Excel.Application: Set ExcelApp = New Excel.Application
    Dim PowerPointApp As PowerPoint.Application: Set PowerPointApp = New PowerPoint.Application

    WordApp.Visible = False
    ExcelApp.Visible = False

    Dim ThemeKey As Variant
    For Each ThemeKey In ThemePaths.Keys
        Dim ThemePath As String: ThemePath = ThemePaths(CStr(ThemeKey))
        Dim ThemeName As String: ThemeName = FileSystem.GetBaseName(ThemePath)
        Debug.Print "Theme name: " & ThemeName

        Dim TargetFolderPath As String
        'TargetFolderPath = App.DownloadsFilePath & "\" & CurrentFolderName & " - themed (" & ThemeName & ")"
        TargetFolderPath = App.DownloadsFilePath & "\Files branded - themed (" & ThemeName & ")"

        If FileSystem.FolderExists(TargetFolderPath) Then FileSystem.DeleteFolder TargetFolderPath, True
        FileSystem.CreateFolder TargetFolderPath

        Me.ApplyThemeToFolderFiles WordApp, ExcelApp, PowerPointApp, ThemePath, SourceFolder, TargetFolderPath, FileSystem
        Shell "explorer.exe """ & TargetFolderPath & """", vbNormalFocus
    Next ThemeKey

    WordApp.Quit: Set WordApp = Nothing
    ExcelApp.Quit: Set ExcelApp = Nothing

End Sub

Private Function IsValidFolderAndThemeSelection( _
    ByVal SourceFolderPath As String, _
    ByVal ThemePaths As Object) As Boolean

    If Dir(SourceFolderPath, vbDirectory) = vbNullString Then
        MsgBox "Selected source folder does not exist:" & vbCrLf & SourceFolderPath, vbCritical
        Exit Function
    End If

    If ThemePaths Is Nothing Then
        MsgBox "No themes in the list to apply.", vbExclamation
        Exit Function
    End If

    If ThemePaths.count = 0 Then
        MsgBox "No themes in the list to apply.", vbExclamation
        Exit Function
    End If

    IsValidFolderAndThemeSelection = True
End Function

