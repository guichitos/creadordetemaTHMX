Attribute VB_Name = "TimerMod"
Option Explicit

Private gTimerID As LongPtr

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" _
        (ByVal hWnd As LongPtr, _
         ByVal nIDEvent As LongPtr, _
         ByVal uElapse As Long, _
         ByVal lpTimerFunc As LongPtr) As LongPtr

    Private Declare PtrSafe Function KillTimer Lib "user32" _
        (ByVal hWnd As LongPtr, _
         ByVal nIDEvent As LongPtr) As Long
#End If

Public Sub StartPolling()
    ' 500 ms es mejor que 10 segundos para detectar click
    gTimerID = SetTimer(0, 0, 500, AddressOf TimerProc)
End Sub

Public Sub StopPolling()
    If gTimerID <> 0 Then
        KillTimer 0, gTimerID
        gTimerID = 0
    End If
End Sub

Public Sub TimerProc( _
    ByVal hWnd As LongPtr, _
    ByVal uMsg As Long, _
    ByVal idEvent As LongPtr, _
    ByVal dwTime As Long)

    If App Is Nothing Then Exit Sub
    If Not App.IsWaitingForLayoutClick Then Exit Sub
    If ActiveWindow Is Nothing Then Exit Sub
    
    If ActiveWindow.ViewType <> ppViewMasterThumbnails Then Exit Sub
    If ActiveWindow.Selection Is Nothing Then Exit Sub
    If ActiveWindow.Selection.Type <> ppSelectionSlides Then Exit Sub
    
    ' Click detectado en layout
    App.IsWaitingForLayoutClick = False
    StopPolling
    
    ' ?? AQUÍ va tu lógica real
    ExecuteDeferredBackgroundCopy
    
End Sub

Public Sub ExecuteDeferredBackgroundCopy()

    MsgBox "Aquí va la copia real del fondo"

    ' Aquí tu lógica real
    ' ApplySlideBackgroundToLayout ...

End Sub
