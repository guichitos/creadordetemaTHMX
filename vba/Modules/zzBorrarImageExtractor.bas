Attribute VB_Name = "zzBorrarImageExtractor"
'Option Explicit
'
'Public Sub ExtractBackgroundImage_Initiator()
'
'    Dim activeSlide As PowerPoint.Slide
'    Set activeSlide = ActiveWindow.View.Slide
'
'    If activeSlide Is Nothing Then Exit Sub
'
'    Dim extractedPath As String
'    extractedPath = ExtractBackgroundImageFromSlide(activeSlide)
'
'    Debug.Print "Extracted background image path: "; extractedPath
'
'End Sub
'
'Public Function ExtractBackgroundImageFromSlide(ByVal slde As PowerPoint.Slide) As String
'
'    ExtractBackgroundImageFromSlide = vbNullString
'
'    If slde Is Nothing Then Exit Function
'    If ActivePresentation.Path = vbNullString Then Exit Function
'
'    Dim PresentationPath As String
'    PresentationPath = GetLocalPath(ActivePresentation.FullName)
'    If PresentationPath = vbNullString Then Exit Function
'
'    Dim FileSystem As Object: Set FileSystem = CreateObject("Scripting.FileSystemObject")
'    Dim ShellApp As Object: Set ShellApp = CreateObject("Shell.Application")
'
'    Dim FileExtension As String
'    FileExtension = LCase$(FileSystem.GetExtensionName(PresentationPath))
'
'    Select Case FileExtension
'        Case "pptx", "pptm", "potx", "potm", "ppsx", "ppsm"
'        Case Else
'            Exit Function
'    End Select
'
'    Dim TempZipPath As String
'    TempZipPath = _
'        FileSystem.GetParentFolderName(PresentationPath) & "\" & _
'        FileSystem.GetBaseName(PresentationPath) & "_temp.zip"
'
'    If FileSystem.FileExists(TempZipPath) Then
'        FileSystem.DeleteFile TempZipPath, True
'    End If
'
'    FileSystem.CopyFile PresentationPath, TempZipPath, True
'
'    Dim ZipFolder As Object
'    Set ZipFolder = ShellApp.Namespace(FileSystem.GetParentFolderName(TempZipPath))
'    If ZipFolder Is Nothing Then GoTo CleanExit
'
'    Dim ZipItem As Object
'    Set ZipItem = ZipFolder.ParseName(FileSystem.GetFileName(TempZipPath))
'    If ZipItem Is Nothing Then GoTo CleanExit
'
'    Dim ZipNamespace As Object
'    Set ZipNamespace = ZipItem.GetFolder
'    If ZipNamespace Is Nothing Then GoTo CleanExit
'
'    Dim SlideIndex As Long
'    SlideIndex = slde.SlideIndex
'
'    Dim SlideXmlPath As String
'    SlideXmlPath = "ppt/slides/slide" & SlideIndex & ".xml"
'
'    Dim SlideRelsPath As String
'    SlideRelsPath = "ppt/slides/_rels/slide" & SlideIndex & ".xml.rels"
'
'    Dim SlideXmlContent As String
'    SlideXmlContent = ReadTextFileFromZip(ZipNamespace, SlideXmlPath)
'    If SlideXmlContent = vbNullString Then GoTo CleanExit
'
'    Dim EmbedId As String
'    EmbedId = ExtractEmbedIdFromSlideXml(SlideXmlContent)
'    If EmbedId = vbNullString Then GoTo CleanExit
'
'    Dim RelsXmlContent As String
'    RelsXmlContent = ReadTextFileFromZip(ZipNamespace, SlideRelsPath)
'    If RelsXmlContent = vbNullString Then GoTo CleanExit
'
'    Dim ImageTarget As String
'    ImageTarget = ExtractImageTargetFromRelsXml(RelsXmlContent, EmbedId)
'    If ImageTarget = vbNullString Then GoTo CleanExit
'
'    Dim ResolvedZipPath As String
'    ResolvedZipPath = "ppt/" & Replace(ImageTarget, "../", "")
'
'    Dim DestinationFolder As String
'    DestinationFolder = FileSystem.GetParentFolderName(PresentationPath)
'
'    CopyFileFromZip ZipNamespace, ResolvedZipPath, DestinationFolder
'
'    ExtractBackgroundImageFromSlide = _
'        DestinationFolder & "\" & FileSystem.GetFileName(ResolvedZipPath)
'
'CleanExit:
'    If FileSystem.FileExists(TempZipPath) Then
'        FileSystem.DeleteFile TempZipPath, True
'    End If
'
'End Function
'
'Private Function ReadTextFileFromZip(ByRef ZipNamespace As Object, ByVal InternalPath As String) As String
'
'    Debug.Print "==== ReadTextFileFromZip START ===="
'
'    If ZipNamespace Is Nothing Then Exit Function
'
'    Dim ShellApp As Object: Set ShellApp = CreateObject("Shell.Application")
'    Dim FileSystem As Object: Set FileSystem = CreateObject("Scripting.FileSystemObject")
'
'    Dim DestinationNamespace As Object
'    Set DestinationNamespace = ShellApp.Namespace(&H5&) ' My Documents
'
'    Debug.Print "DestinationNamespace Is Nothing: "; DestinationNamespace Is Nothing
'    If DestinationNamespace Is Nothing Then
'        Debug.Print "FATAL: Cannot access Documents folder via Shell."
'        Exit Function
'    End If
'
'    Dim PathParts() As String
'    PathParts = Split(Replace(InternalPath, "/", "\"), "\")
'
'    Dim CurrentFolder As Object
'    Set CurrentFolder = ZipNamespace
'
'    Dim i As Long
'    For i = LBound(PathParts) To UBound(PathParts) - 1
'        Dim FolderItem As Object
'        Set FolderItem = CurrentFolder.ParseName(PathParts(i))
'        If FolderItem Is Nothing Then Exit Function
'
'        Set CurrentFolder = FolderItem.GetFolder
'        If CurrentFolder Is Nothing Then Exit Function
'    Next i
'
'    Dim FileItem As Object
'    Set FileItem = CurrentFolder.ParseName(PathParts(UBound(PathParts)))
'    If FileItem Is Nothing Then Exit Function
'
'    DestinationNamespace.CopyHere FileItem, 16
'
'    Dim TempFilePath As String
'    TempFilePath = DestinationNamespace.Self.Path & "\" & FileItem.Name
'
'    Dim StartTime As Single: StartTime = Timer
'    Do While Not FileSystem.FileExists(TempFilePath)
'        If Timer - StartTime > 5 Then Exit Function
'        DoEvents
'    Loop
'
'    Dim TextStream As Object
'    Set TextStream = FileSystem.OpenTextFile(TempFilePath, 1)
'
'    ReadTextFileFromZip = TextStream.ReadAll
'    TextStream.Close
'
'    FileSystem.DeleteFile TempFilePath, True
'
'    Debug.Print "Read length: "; Len(ReadTextFileFromZip)
'    Debug.Print "==== ReadTextFileFromZip END (SUCCESS) ===="
'
'End Function
'
'Private Function ExtractEmbedIdFromSlideXml(ByVal SlideXml As String) As String
'
'    Dim startPos As Long
'    startPos = InStr(SlideXml, "r:embed=""")
'    If startPos = 0 Then Exit Function
'
'    startPos = startPos + 9
'    Dim EndPos As Long
'    EndPos = InStr(startPos, SlideXml, """")
'
'    ExtractEmbedIdFromSlideXml = Mid(SlideXml, startPos, EndPos - startPos)
'
'End Function
'Private Function ExtractImageTargetFromRelsXml(ByVal RelsXml As String, ByVal EmbedId As String) As String
'
'    Dim SearchToken As String
'    SearchToken = "Id=""" & EmbedId & """"
'
'    Dim pos As Long
'    pos = InStr(RelsXml, SearchToken)
'    If pos = 0 Then Exit Function
'
'    pos = InStr(pos, RelsXml, "Target=""")
'    If pos = 0 Then Exit Function
'
'    pos = pos + 8
'
'    Dim EndPos As Long
'    EndPos = InStr(pos, RelsXml, """")
'
'    If EndPos = 0 Then Exit Function
'
'    ExtractImageTargetFromRelsXml = Mid(RelsXml, pos, EndPos - pos)
'
'End Function
'
'Private Sub CopyFileFromZip(ByRef ZipNamespace As Object, ByVal InternalPath As String, ByVal DestinationFolder As String)
'
'    Debug.Print "==== CopyFileFromZip START ===="
'    Debug.Print "InternalPath expected: "; InternalPath
'    Debug.Print "Final destination folder (after copy): "; DestinationFolder
'
'    Dim ShellApp As Object: Set ShellApp = CreateObject("Shell.Application")
'    Dim FileSystem As Object: Set FileSystem = CreateObject("Scripting.FileSystemObject")
'
'    Dim ShellDestination As Object
'    Set ShellDestination = ShellApp.Namespace(&H5&) ' My Documents
'
'    Debug.Print "ShellDestination (Documents) Is Nothing: "; ShellDestination Is Nothing
'    If ShellDestination Is Nothing Then
'        Debug.Print "FATAL: Cannot access Documents as Shell namespace."
'        Debug.Print "==== CopyFileFromZip END (FAIL) ===="
'        Exit Sub
'    End If
'
'    Dim PathParts() As String
'    PathParts = Split(Replace(InternalPath, "/", "\"), "\")
'
'    Dim CurrentFolder As Object
'    Set CurrentFolder = ZipNamespace
'
'    Dim i As Long
'    For i = LBound(PathParts) To UBound(PathParts) - 1
'
'        Debug.Print "Entering folder: "; PathParts(i)
'
'        Dim FolderItem As Object
'        Set FolderItem = CurrentFolder.ParseName(PathParts(i))
'
'        Debug.Print "FolderItem Is Nothing: "; FolderItem Is Nothing
'        If FolderItem Is Nothing Then
'            Debug.Print "EXIT: Folder not found in ZIP: "; PathParts(i)
'            Debug.Print "==== CopyFileFromZip END (FAIL) ===="
'            Exit Sub
'        End If
'
'        Set CurrentFolder = FolderItem.GetFolder
'        If CurrentFolder Is Nothing Then
'            Debug.Print "EXIT: Cannot open ZIP folder: "; PathParts(i)
'            Debug.Print "==== CopyFileFromZip END (FAIL) ===="
'            Exit Sub
'        End If
'
'        Debug.Print "Current ZIP folder path: "; CurrentFolder.Self.Path
'    Next i
'
'    Dim fileName As String
'    fileName = PathParts(UBound(PathParts))
'    Debug.Print "Looking for file: "; fileName
'
'    Dim FileItem As Object
'    Set FileItem = CurrentFolder.ParseName(fileName)
'
'    Debug.Print "FileItem Is Nothing: "; FileItem Is Nothing
'    If FileItem Is Nothing Then
'        Debug.Print "EXIT: File not found in ZIP."
'        Debug.Print "==== CopyFileFromZip END (FAIL) ===="
'        Exit Sub
'    End If
'
'    Debug.Print "ZIP file found: "; FileItem.Path
'    Debug.Print "Copying file to Documents via Shell..."
'
'    ShellDestination.CopyHere FileItem, 16
'
'    Dim TempCopiedPath As String
'    TempCopiedPath = ShellDestination.Self.Path & "\" & FileItem.Name
'
'    Dim StartTime As Single: StartTime = Timer
'    Do While Not FileSystem.FileExists(TempCopiedPath)
'        If Timer - StartTime > 5 Then
'            Debug.Print "EXIT: Timeout waiting for Shell copy."
'            Debug.Print "==== CopyFileFromZip END (FAIL) ===="
'            Exit Sub
'        End If
'        DoEvents
'    Loop
'
'    Debug.Print "File copied to Documents: "; TempCopiedPath
'
'    If DestinationFolder <> ShellDestination.Self.Path Then
'        Debug.Print "Moving file to final destination..."
'
'        If Not FileSystem.FolderExists(DestinationFolder) Then
'            FileSystem.CreateFolder DestinationFolder
'        End If
'
'        Dim FinalPath As String
'        FinalPath = DestinationFolder & "\" & FileItem.Name
'
'        If FileSystem.FileExists(FinalPath) Then
'            Debug.Print "Destination file exists, deleting: "; FinalPath
'            FileSystem.DeleteFile FinalPath, True
'        End If
'
'        FileSystem.MoveFile TempCopiedPath, FinalPath
'        Debug.Print "File moved to: "; FinalPath
'    End If
'
'    Debug.Print "==== CopyFileFromZip END (SUCCESS) ===="
'
'End Sub
'
