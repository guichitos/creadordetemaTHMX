Attribute VB_Name = "VbaRepoBridge"
'==========================
' Module: VbaRepoBridge
' Purpose: Safe export of VBA components to filesystem
'==========================

Option Explicit


Public Sub ExportVbaProjectToRepo()
    On Error GoTo CleanFail

    Dim exportRoot As String
    exportRoot = GetExportRootPath()

    EnsureFolderExists exportRoot
    EnsureFolderExists exportRoot & "\vba"
    EnsureFolderExists exportRoot & "\vba\Modules"
    EnsureFolderExists exportRoot & "\vba\Classes"
    EnsureFolderExists exportRoot & "\vba\Forms"

    Dim vbComp As Object
    For Each vbComp In Application.VBE.VBProjects(1).VBComponents
        ExportComponent vbComp, exportRoot
    Next vbComp

    MsgBox "VBA export completed successfully." & vbCrLf & exportRoot, vbInformation
    Shell "explorer.exe """ & exportRoot & """", vbNormalFocus

    Exit Sub

CleanFail:
    MsgBox "Export failed: " & Err.Description, vbCritical
End Sub

'--------------------------
' Helpers
'--------------------------

Private Sub ExportComponent(vbComp As Object, rootPath As String)
    Dim targetFolder As String

    Select Case vbComp.Type
        Case 1 ' vbext_ct_StdModule
            targetFolder = rootPath & "\vba\Modules"

        Case 2 ' vbext_ct_ClassModule
            targetFolder = rootPath & "\vba\Classes"
        Case 3 ' vbext_ct_MSForm
            targetFolder = rootPath & "\vba\Forms"
        Case Else
            Exit Sub
    End Select

    Dim FilePath As String
    FilePath = targetFolder & "\" & vbComp.Name & GetExtension(vbComp.Type)

    On Error Resume Next
    Kill FilePath
    On Error GoTo 0

    vbComp.Export FilePath
End Sub

Private Function GetExtension(vbCompType As Long) As String
    Select Case vbCompType
        Case 1: GetExtension = ".bas"
        Case 2: GetExtension = ".cls"
        Case 3: GetExtension = ".frm"
        Case Else: GetExtension = ""
    End Select
End Function

Private Function GetExportRootPath() As String
    Dim documentsPath As String
    documentsPath = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")

    GetExportRootPath = documentsPath & "\Thmx-ppt-vba-project"
End Function


Private Sub EnsureFolderExists(folderPath As String)
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
End Sub

'
'| Módulo                    | Tipo   | Uso |
'|---------------------------|--------|-----|
'| aatest                    | Module | Módulo de prueba vacío usado como placeholder. |
'| App                       | Class  | Centraliza el acceso al host de PowerPoint y expone objetos clave como comandos, barras y usuario activo. |
'| AppCommandBars            | Class  | Construye y actualiza menús contextuales y CommandBars con controles asociados a acciones. |
'| AppCommands               | Class  | Agrupa los comandos del add-in para formatear tablas/gráficos, ajustar vistas y manipular formas. |
'| AppCommandsRibbon         | Module | Gestiona comandos de la cinta para guardar y aplicar el zoom basado en un área de trabajo en la diapositiva. |
'| AppCommandsUtilities      | Module | Define enumeraciones de colores, estilos y tipos de gráficos para estandarizar comandos relacionados con gráficos. |
'| AppContextMenus           | Module | Centraliza las constantes y etiquetas usadas por los menús contextuales y sus botones. |
'| AppCustomUiUtilities      | Module | Administra la carga, validación y control de la interfaz Ribbon/CustomUI en Office. |
'| AppDefaultCustomUi        | Class  | Implementa la configuración por defecto del Ribbon y sus callbacks para la interfaz personalizada. |
'| AppErrors                 | Module | Provee utilidades para lanzar, registrar y guardar errores en archivo de log. |
'| AppMacros                 | Module | Expone macros públicas que conectan acciones de menús con la lógica de App.User. |
'| AppProgressBar            | Class  | Gestiona la barra de progreso usada por el add-in para mostrar avance de procesos. |
'| AppStrings                | Module | Gestiona traducciones y mensajes de texto según el idioma de la aplicación. |
'| AppUser                   | Class  | Centraliza el estado y las acciones del usuario sobre selecciones y operaciones del add-in. |
'| AppUtilities              | Module | Inicializa objetos principales y expone el singleton App y utilidades base de la aplicación. |
'| DeveloperUtilities        | Module | Incluye utilidades de desarrollo para inspeccionar y limpiar CommandBars y etiquetas. |
'| FileSystem                | Class  | Abstrae operaciones de sistema de archivos y conversiones de rutas para el add-in. |
'| GoogleAuthenticator       | Class  | Gestiona autenticación para servicios Google usados por el add-in. |
'| GoogleTranslationApi      | Class  | Envuelve llamadas a la API de traducción de Google. |
'| IAppCustomUi              | Class  | Define la interfaz que implementan las clases de Custom UI del Ribbon. |
'| IWebAuthenticator         | Class  | Define la interfaz para autenticadores usados en clientes web. |
'| Learnings                 | Module | Contiene fragmentos de aprendizaje y pruebas internas con objetos de PowerPoint. |
'| PowerPointApplication     | Class  | Envuelve la aplicación de PowerPoint para exponer helpers y estado. |
'| PowerPointApplicationMode | Class  | Representa modos/configuraciones de la aplicación PowerPoint. |
'| PowerPointPresentation    | Class  | Encapsula operaciones y datos de una presentación de PowerPoint. |
'| PowerPointPresentations   | Class  | Gestiona colecciones de presentaciones abiertas en PowerPoint. |
'| PowerPointUtilities       | Module | Ofrece utilidades generales de PowerPoint como propiedades, selección y visibilidad de paneles. |
'| PowerPointVbProject       | Class  | Maneja operaciones relacionadas con el proyecto VBA de PowerPoint. |
'| ShapeGroup                | Class  | Representa un grupo de formas con metadatos y operaciones asociadas. |
'| ShapeGroupUtilities       | Module | Administra el registro y consulta de grupos de formas identificados en la presentación. |
'| SlideTransferManager      | Class  | Coordina la transferencia y selección de diapositivas entre presentaciones. |
'| SlideTransferUtilities    | Module | Provee funciones auxiliares para seleccionar, validar y transferir diapositivas con miniaturas. |
'| Speech                    | Module | Implementa síntesis de voz para leer texto mediante SAPI. |
'| StringsUtilities          | Module | Contiene utilidades para limpiar cadenas y eliminar caracteres ocultos. |
'| VbaRepoBridge             | Module | Exporta componentes VBA al sistema de archivos para sincronizarlos con el repositorio. |
'| VbaUtilities              | Module | Agrupa utilidades VBA generales como carga de imágenes y apertura de URLs. |
'| WebApiFetch               | Class  | Implementa solicitudes HTTP y su configuración para consumir APIs web. |
'| WebApiFetchTests          | Class  | Incluye pruebas o helpers de prueba para WebApiFetch. |
'| WebClient                 | Class  | Cliente de alto nivel para realizar peticiones web y manejar respuestas. |
'| WebHelpers                | Module | Biblioteca de apoyo para solicitudes web, conversiones y utilidades de VBA-Web. |
'| WebRequest                | Class  | Representa una solicitud web con parámetros, headers y cuerpo. |
'| WebResponse               | Class  | Representa la respuesta de una solicitud web y sus datos asociados. |
'



'Secuencia de módulos/procesos
'1. Cargar/actualizar los menús contextuales
'
    'Al cargar el Ribbon, AppCustomUiUtilities.onLoad invoca App.CommandBars.Update, lo que fuerza la construcción/actualización de los menús contextuales desde la app.
    'Además, justo antes del clic derecho, el evento pPowerPointApp_WindowBeforeRightClick en AppUser vuelve a llamar App.CommandBars.Update, asegurando que el menú esté actualizado cuando aparece.

'2. Construcción del menú de thumbnails
    '
    'AppCommandBars.Update llama a BuildContextMenus y valida que el menú de thumbnails exista (ThumbnailsBuiltInContextMenu).
    'ThumbnailsBuiltInContextMenu toma el CommandBar integrado de PowerPoint "Thumbnails" y agrega controles personalizados, incluyendo “Create layout”.
    '
'3. Inserción del botón “Create layout”
'
    'La inserción se hace con AddCustomControlTo, que crea el control si no existe y le asigna OnAction = "AppContextMenus.ControlActioned".
    'Para CREATE_LAYOUT_BUTTON_ID, el botón se hace visible y habilitado solo cuando la vista de thumbnails está activa y hay selección de slides.
'
'4.Ejecución de la acción al hacer clic
'
    'Al hacer clic, AppContextMenus.ControlActioned inspecciona el Tag del control. Si es CREATE_LAYOUT_BUTTON_ID, delega a App.User.WantsToCreateALayoutBasedOnCurrentSlide.
'
'5. Validación previa y selección de design
'
    'WantsToCreateALayoutBasedOnCurrentSlide revisa si hay shapes decorativas; si no las hay, pregunta confirmación al usuario. Luego muestra el menú SelectSlideMasterContextMenu para elegir el Design (o crear uno nuevo).
    'SelectSlideMasterContextMenu se construye con botones por cada Design y un textbox para un nuevo nombre, también apuntando a AppContextMenus.ControlActioned.
'
'6. Creación final del layout
'
    'Al elegir un Design (o escribir uno nuevo), ControlActioned llama a WantsToCreateALayoutBasedOnCurrentSlideOnADesignCalled.
    'Ese método invoca App.Commands.CreateLayoutFrom, que crea el layout a partir del slide actual (vía CreateCustomLayoutFromSlide) y limpia placeholders.
    
'
'**Resumen compacto del flujo (modo “pipeline”)
    '1. onLoad o WindowBeforeRightClick ? App.CommandBars.Update
'
    '2. Update ? ThumbnailsBuiltInContextMenu ? agrega Create layout en CommandBars("Thumbnails")
'
    '3. Click en el botón ? AppContextMenus.ControlActioned ? WantsToCreateALayoutBasedOnCurrentSlide
'
    '4. Confirma, muestra menú SelectSlideMasterContextMenu ? selecciona Design
'
    '5. CreateLayoutFrom genera el CustomLayout y lo limpia
    
    
'************Menu contextual

'Resumen del flujo (“Match width to the widest”)
'Creación del menú/botón en el contexto “Arrange > Sizing”
'En AppCommandBars.cls, dentro del menú de “Arrange” se crea el submenú “Sizing” y allí se agrega el botón con caption “Match width to the widest”. Este botón configura OnAction a AppContextMenus.ControlActioned y se le asigna el tag MATCH_MAX_WIDTH_TAG, que es lo que dispara el flujo posterior.
'
'Despacho del evento al presionar el botón
'En AppContextMenus.ControlActioned, se lee el ActionControl.Tag y se compara en el Select Case. Cuando el tag coincide con MATCH_MAX_WIDTH_TAG, se llama a ThisApp.User.WantsToMatchMaxWidth, que es el método real que ejecuta la lógica de ajuste de ancho.
'
'Lógica aplicada al presionar “Match width to the widest”
'WantsToMatchMaxWidth (en AppUser.cls) valida que la selección sea de shapes; luego calcula el ancho máximo entre las shapes seleccionadas y finalmente asigna ese ancho máximo a todas las shapes seleccionadas. Esto asegura que todas queden con el ancho de la más ancha.
'

