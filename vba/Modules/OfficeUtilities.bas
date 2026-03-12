Attribute VB_Name = "OfficeUtilities"
Option Explicit

Public Enum ColorPlacement
    ToLight1
    ToDark1
    ToLight2
    ToDark2
    
    ToAccent1
    ToAccent2
    ToAccent3
    ToAccent4
    ToAccent5
    ToAccent6
    
    ToHiperlink
    ToVisitedHyperlink
End Enum

Public Type ColorMapping
    Background1 As ColorPlacement
    Text1 As ColorPlacement
    Background2 As ColorPlacement
    Text2 As ColorPlacement
    
    Accent1 As ColorPlacement
    Accent2 As ColorPlacement
    Accent3 As ColorPlacement
    Accent4 As ColorPlacement
    Accent5 As ColorPlacement
    Accent6 As ColorPlacement
    
    Hiperlink As ColorPlacement
    VisitedHyperlink As ColorPlacement
End Type

Public Type OfficeColorPalette
    Name As String
    
    Light1 As Long
    Dark1 As Long
    Light2 As Long
    Dark2 As Long
    
    Accent1 As Long
    Accent2 As Long
    Accent3 As Long
    Accent4 As Long
    Accent5 As Long
    Accent6 As Long
    
    Hyperlink As Long
    VisitedHyperlink As Long
    
    Mapping As ColorMapping
End Type

Public Function GetColorPalette(ByVal Name As String, _
                                    Light1 As Long, Dark1 As Long, _
                                    Light2 As Long, Dark2 As Long, _
                                    Accent1 As Long, Accent2 As Long, Accent3 As Long, Accent4 As Long, Accent5 As Long, Accent6 As Long, _
                                    Hyperlink As Long, VisitedHyperlink As Long, _
                                    Mapping As ColorMapping) As OfficeColorPalette
    Dim Result As OfficeColorPalette
    With Result
        .Name = Name
        
        .Light1 = Light1
        .Dark1 = Dark1
        .Light2 = Light2
        .Dark2 = Dark2
        
        .Accent1 = Accent1
        .Accent2 = Accent2
        .Accent3 = Accent3
        .Accent4 = Accent4
        .Accent5 = Accent5
        .Accent6 = Accent6
        
        .Hyperlink = Hyperlink
        .VisitedHyperlink = VisitedHyperlink
        
        .Mapping = Mapping
    End With
    
    GetColorPalette = Result
End Function

Public Function GetMapping(Optional ByVal Background1 As ColorPlacement = ToLight1, _
                            Optional ByVal Text1 As ColorPlacement = ToDark1, _
                            Optional ByVal Background2 As ColorPlacement = ToLight2, _
                            Optional ByVal Text2 As ColorPlacement = ToDark2, _
                            Optional ByVal Accent1 As ColorPlacement = ToAccent1, _
                            Optional ByVal Accent2 As ColorPlacement = ToAccent2, _
                            Optional ByVal Accent3 As ColorPlacement = ToAccent3, _
                            Optional ByVal Accent4 As ColorPlacement = ToAccent4, _
                            Optional ByVal Accent5 As ColorPlacement = ToAccent5, _
                            Optional ByVal Accent6 As ColorPlacement = ToAccent6, _
                            Optional ByVal Hiperlink As ColorPlacement = ToHiperlink, _
                            Optional ByVal VisitedHyperlink As ColorPlacement = ToVisitedHyperlink) As ColorMapping
    With GetMapping
        .Background1 = Background1
        .Text1 = Text1
        .Background2 = Background2
        .Text2 = Text2
        
        .Accent1 = Accent1
        .Accent2 = Accent2
        .Accent3 = Accent3
        .Accent4 = Accent4
        .Accent5 = Accent5
        .Accent6 = Accent6
        
        .Hiperlink = Hiperlink
        .VisitedHyperlink = VisitedHyperlink
    End With
End Function

Public Function DefaultColorMapping() As ColorMapping
    DefaultColorMapping = GetMapping
End Function

Public Function ColorPlacementToString(ByVal Placement As ColorPlacement) As String
    Dim Result As String
    Select Case True
        Case Placement = ToLight1:              Result = "lt1"
        Case Placement = ToDark1:               Result = "dk1"
        Case Placement = ToLight2:              Result = "lt2"
        Case Placement = ToDark2:               Result = "dk2"
        
        Case Placement = ToAccent1:             Result = "accent1"
        Case Placement = ToAccent2:             Result = "accent2"
        Case Placement = ToAccent3:             Result = "accent3"
        Case Placement = ToAccent4:             Result = "accent4"
        Case Placement = ToAccent5:             Result = "accent5"
        Case Placement = ToAccent6:             Result = "accent6"
        
        Case Placement = ToHiperlink:           Result = "hlink"
        Case Placement = ToVisitedHyperlink:    Result = "folHlink"
    End Select
    ColorPlacementToString = Result
End Function

