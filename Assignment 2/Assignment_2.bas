Attribute VB_Name = "Module1"
Option Explicit

Sub AddNumbersA()
Dim num As Double
    num = InputBox("Nh?p m?t s?:")
    Range("D4").Value = Range("D4").Value + num
    Range("G12").Value = Range("D4").Value
End Sub

Sub AddNumbersB()
Dim num As Double
    num = InputBox("Nh?p m?t s?:")
    ActiveCell.Value = ActiveCell.Value + num
    ActiveCell.Offset(-3, 2).Value = ActiveCell.Value
End Sub

Sub WherePutMe()
Dim rowNum As Integer
    Dim colLetter As String
    rowNum = InputBox("Nh?p s? hàng:")
    colLetter = InputBox("Nh?p ch? cái c?a c?t:")
    Range(colLetter & rowNum).Value = Selection.Cells(2, 2).Value
End Sub

Sub Swap()
 Dim Temp As Variant
    Temp = Selection.Cells(1, 1).Value
    Selection.Cells(1, 1).Value = Selection.Cells(1, 2).Value
    Selection.Cells(1, 2).Value = Temp
End Sub
