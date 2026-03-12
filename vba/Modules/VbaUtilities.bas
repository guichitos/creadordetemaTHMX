Attribute VB_Name = "VbaUtilities"
Option Explicit

Public Function Min(ByVal arr As Variant) As Variant
    Min = arr(LBound(arr))
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) < Min Then Min = arr(i)
    Next i
End Function

Public Function Max(ByVal arr As Variant) As Variant
    Max = arr(LBound(arr))
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) > Max Then Max = arr(i)
    Next i
End Function
