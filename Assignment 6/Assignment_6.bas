Attribute VB_Name = "Module1"
Option Explicit

Sub Cookies()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Ð?t tên cho sheet b?n dang làm vi?c
    Set ws = ThisWorkbook.Sheets("February")
    
    ' Tìm hàng cu?i cùng có d? li?u trong c?t A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Thi?t l?p công th?c SUMIF cho c?t I t? I6 d?n I10
    ws.Range("I6:I10").FormulaR1C1 = "=SUMIFS(R6C6:R" & lastRow & "C6, R6C1:R" & lastRow & "C1, RC[-1])"
    
    ' Thi?t l?p công th?c SUM cho c?t I11
    ws.Range("I11").FormulaR1C1 = "=SUM(R6C6:R" & lastRow & "C6)"
    
    ' Thi?t l?p công th?c SUMIF cho c?t I t? I6 d?n I10
    ws.Range("I14:I16").FormulaR1C1 = "=SUMIFS(R6C4:R" & lastRow & "C4, R6C3:R" & lastRow & "C3, RC[-1])"

End Sub



Sub Reset()
'Do not modify or delete this sub
Sheets("Original Data").Columns("A:F").Copy Sheets("February").Columns("A:F")
With Sheets("February")
    .Range("I6:I11").ClearContents
    .Range("I14:I16").ClearContents
End With
End Sub
