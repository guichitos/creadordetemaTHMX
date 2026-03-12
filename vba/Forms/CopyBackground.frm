VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CopyBackground 
   Caption         =   "Copy Background"
   ClientHeight    =   2205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4260
   OleObjectBlob   =   "CopyBackground.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CopyBackground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Unload Me
    App.IsWaitingForLayoutClick = False
    On Error Resume Next
    Application.CommandBars.ExecuteMso "MasterViewClose"
    On Error GoTo 0
    'SyncLayoutsFromSampleSlidesShapesAndPlaceholders
    App.ThemePreviewer.PreviewTheme
End Sub

