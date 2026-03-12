Attribute VB_Name = "AppMacros"
Option Explicit

Public App As App
Private ThemePreviewer As AppThemePreviewer
Private ThemeCreator As AppThemeCreator
Private ThemeApplier As AppThemeApplier
Public g_Ribbon As IRibbonUI

Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set g_Ribbon = ribbon
End Sub

Public Sub LoadApp()
    EnsureAppInitialized
End Sub

Public Sub UnloadApp()
    Set App = Nothing
End Sub

Sub PreviewApp(control As IRibbonControl)
    EnsureAppInitialized
    If ThemePreviewer Is Nothing Then Set ThemePreviewer = New AppThemePreviewer
    ThemePreviewer.PreviewTheme
    If App.IsWaitingForLayoutClick = True Then Exit Sub
End Sub

Sub CreateApp(control As IRibbonControl)
    EnsureAppInitialized
    If ThemeCreator Is Nothing Then Set ThemeCreator = New AppThemeCreator
    ThemeCreator.CreateTheme True
End Sub

Sub ApplyApp(control As IRibbonControl)
    
    EnsureAppInitialized
    If ThemeApplier Is Nothing Then Set ThemeApplier = New AppThemeApplier
    ThemeApplier.ApplyTheme
End Sub

Sub GetChkBackgroundPressed(control As IRibbonControl, ByRef returnedVal)
    If App Is Nothing Then
        returnedVal = False
        Exit Sub
    End If

    returnedVal = App.ShouldIgnoreComplexBackgrounds
End Sub

Sub ChkBackgrounds_Click(control As IRibbonControl, pressed As Boolean)

    EnsureAppInitialized
    App.ShouldIgnoreComplexBackgrounds = pressed
    App.CurrentSlideIndex = 0
    App.IsWaitingForLayoutClick = False
    
    If Not g_Ribbon Is Nothing Then
        g_Ribbon.InvalidateControl control.Id
    End If

End Sub
Sub RefreshRibbon()

    If g_Ribbon Is Nothing Then Exit Sub

    g_Ribbon.Invalidate

End Sub
Sub getEnabled(control As IRibbonControl, ByRef returnedVal)
    If App Is Nothing Then
        returnedVal = False
        Exit Sub
    End If

    returnedVal = App.LayoutDesignHasComplexBackgrounds
End Sub

Private Sub EnsureAppInitialized()
    If Not App Is Nothing Then Exit Sub

    Set App = New App
    App.ShouldIgnoreComplexBackgrounds = False
    App.IsWaitingForLayoutClick = False
    App.CurrentSlideIndex = 0
End Sub
