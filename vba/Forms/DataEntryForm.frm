VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataEntryForm 
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6045
   OleObjectBlob   =   "DataEntryForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "DataEntryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ThemePaths As Object

Private Sub RemoveThemeToListButton_Click()
    Dim SelectedIndex As Long
    SelectedIndex = ThemeListbox.ListIndex

    If SelectedIndex = -1 Then
        MsgBox "Please select a theme to remove.", vbExclamation
        Exit Sub
    End If

    Dim ThemeName As String
    ThemeName = ThemeListbox.List(SelectedIndex)

    If ThemePaths.Exists(ThemeName) Then
        ThemePaths.Remove ThemeName
    End If

    ThemeListbox.RemoveItem SelectedIndex
End Sub

Private Sub UserForm_Initialize()
    Set ThemePaths = CreateObject("Scripting.Dictionary")
    Dim BasePath As String

    BasePath = ActivePresentation.Path
    DataEntryForm.BusinessAddressStreetTextBox.value = GetLocalPath(BasePath & "\Files to be branded")
    LoadLastGeneratedThemes
    
End Sub

Private Sub LoadLastGeneratedThemes()
        
        If App.ThemeCreator.GetBrandThemePath <> "" Then
            ThemeListbox.AddItem CreateObject("Scripting.FileSystemObject").GetFileName(App.ThemeCreator.GetBrandThemePath)
            ThemePaths.Add CreateObject("Scripting.FileSystemObject").GetFileName(App.ThemeCreator.GetBrandThemePath), App.ThemeCreator.GetBrandThemePath
        End If
        
        If App.ThemeCreator.GetUnreliableThemePath <> "" Then
            ThemeListbox.AddItem CreateObject("Scripting.FileSystemObject").GetFileName(App.ThemeCreator.GetUnreliableThemePath)
            ThemePaths.Add CreateObject("Scripting.FileSystemObject").GetFileName(App.ThemeCreator.GetUnreliableThemePath), App.ThemeCreator.GetUnreliableThemePath
        End If
End Sub

Private Sub ApplyThemesButton_Click()
    ApplyThemesToFolderContent
    Unload Me
End Sub

Private Sub SelectFolderButton_Click()
    Dim FolderDialog As FileDialog
    Dim SelectedFolderPath As String

    Set FolderDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With FolderDialog
        .Title = "Select a folder"
        If .Show <> -1 Then Exit Sub
        SelectedFolderPath = .SelectedItems(1)
    End With

    BusinessAddressStreetTextBox.value = SelectedFolderPath
End Sub

Private Function IsThemeAlreadyListed(ByVal ThemePath As String) As Boolean
    Dim Index As Long

    For Index = 0 To ThemeListbox.ListCount - 1
        If ThemeListbox.List(Index) = ThemePath Then
            IsThemeAlreadyListed = True
            Exit Function
        End If
    Next
End Function

Private Sub AddThemeToListButton_Click()
    Dim FileDialog As FileDialog
    Dim ThemePath As String
    Dim ThemeName As String

    Set FileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With FileDialog
        .Title = "Select a theme file (.thmx)"
        .Filters.Clear
        .Filters.Add "Office Themes", "*.thmx"
        If .Show <> -1 Then Exit Sub
        ThemePath = .SelectedItems(1)
    End With

    ThemeName = CreateObject("Scripting.FileSystemObject").GetFileName(ThemePath)

    If Not ThemePaths.Exists(ThemeName) Then
        ThemeListbox.AddItem ThemeName
        ThemePaths.Add ThemeName, ThemePath
    End If
End Sub

Private Sub ApplyThemesToFolderContent()
    Dim SourceFolderPath As String
    SourceFolderPath = Trim(DataEntryForm.BusinessAddressStreetTextBox.value)
    If Len(SourceFolderPath) = 0 Then
        CreateTemplatesForThemesInSlectionControl
    Else
        App.ThemeApplier.ApplyThemesToFolder SourceFolderPath, ThemePaths
    End If
End Sub

Private Sub CreateTemplatesForThemesInSlectionControl()
    Dim FileSystem As Object: Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Dim ThemeIndex As Long
    For ThemeIndex = 0 To DataEntryForm.ThemeListbox.ListCount - 1
        Dim ThemeKey As String: ThemeKey = DataEntryForm.ThemeListbox.List(ThemeIndex)
        Dim ThemePath As String: ThemePath = ThemePaths(ThemeKey)
        Dim ThemeName As String: ThemeName = FileSystem.GetBaseName(ThemePath)
        Debug.Print "Theme name: " & ThemeName
        CreateTemplatesWithThemesApplied (ThemePath)
    Next

    Unload Me
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    
    Dim DownloadsPath As String
    DownloadsPath = WshShell.ExpandEnvironmentStrings("%USERPROFILE%\Downloads")
           
    Shell "explorer.exe """ & DownloadsPath & """", vbNormalFocus
End Sub

Private Sub CreateTemplatesWithThemesApplied(ByVal ThemePath As String)
    Debug.Print "ThemePath - " & ThemePath
    Dim FileSystem As Object: Set FileSystem = CreateObject("Scripting.FileSystemObject")
    Dim ThemeFileName As String: ThemeFileName = FileSystem.GetBaseName(ThemePath)
    Debug.Print "ThemeFileName - " & ThemeFileName
    Dim ContainingFolderPath As String: ContainingFolderPath = FileSystem.GetParentFolderName(ThemePath)
    Debug.Print "ContainingFolderPath - " & ContainingFolderPath
    Dim ParentFolderPath As String: ParentFolderPath = FileSystem.GetParentFolderName(ContainingFolderPath)
    Debug.Print "ParentFolderPath - " & ParentFolderPath
    Dim TemplatesFolderPath As String
    Debug.Print "TemplatesFolderPath - " & TemplatesFolderPath
    TemplatesFolderPath = ParentFolderPath & "\Templates for " & ThemeFileName

    CreateFolder TemplatesFolderPath
    App.ThemeApplier.CreateTemplatesFromThemeIn ThemePath, TemplatesFolderPath
    Shell "explorer.exe """ & TemplatesFolderPath & """", vbNormalFocus
    
End Sub


