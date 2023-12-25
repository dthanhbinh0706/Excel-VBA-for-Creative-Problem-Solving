Attribute VB_Name = "Module1"
Option Explicit

Sub FormatAndIncompleteOrders()
    Dim wsNewOrders As Worksheet
    Dim wsIncompleteOrders As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' �?t t�n cho c�c worksheet
    Set wsNewOrders = ThisWorkbook.Sheets("New Orders")
    Set wsIncompleteOrders = ThisWorkbook.Sheets("Incomplete Orders")

    ' T�m h�ng cu?i c�ng c� d? li?u trong c?t D c?a sheet "New Orders"
    lastRow = wsNewOrders.Cells(wsNewOrders.Rows.Count, "D").End(xlUp).Row

    ' 1. X�a to�n b? format tr�n v�ng d? li?u t? h�ng 4 d?n h�ng cu?i c�ng c?a c?t A d?n D
    wsNewOrders.Range("A4:D" & lastRow).Style = "Normal"
    With wsNewOrders.Range("A4:A" & lastRow)
        .NumberFormat = "dd/mmm/yyyy"  ' �?nh d?ng ng�y th�ng theo � mu?n
    End With

    ' 2. L?c v� x�a c�c d�ng c� gi� tr? tr?ng trong c?t B
    With wsNewOrders.Range("A3:D" & lastRow)
        .AutoFilter Field:=2, Criteria1:=""
        .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).SpecialCells(xlCellTypeVisible).ClearContents
    End With

    ' B? l?c
    wsNewOrders.AutoFilterMode = False
    
    ' Di chuy?n d? li?u l�n n?u c� h�ng tr?ng
    For i = lastRow - 1 To 3 Step -1
        If IsEmpty(wsNewOrders.Cells(i, 2)) Then
            wsNewOrders.Rows(i).Delete
            wsNewOrders.Rows(i + 1).Copy wsNewOrders.Rows(i)
        End If
    Next i
    

    ' 3. L?c v� sao ch�p c�c d�ng c� gi� tr? tr?ng trong c?t C sang sheet "Incomplete Orders"
     With wsNewOrders.Range("A3:D" & lastRow)
        .AutoFilter Field:=3, Criteria1:="="
        ' Copy header
        .Resize(1, .Columns.Count).Copy wsIncompleteOrders.Cells(1, 1)
        ' Copy d? li?u (kh�ng bao g?m header)
        .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).SpecialCells(xlCellTypeVisible).Copy wsIncompleteOrders.Cells(wsIncompleteOrders.Rows.Count, 1).End(xlUp).Offset(1, 0)
        ' X�a c�c h�ng d? li?u v?a du?c l?c ? sheet "New Orders"
        .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).SpecialCells(xlCellTypeVisible).EntireRow.Delete
    End With

    ' B? l?c
    wsNewOrders.AutoFilterMode = False

End Sub

Sub Report()
Dim wsNewOrders As Worksheet
    Dim wsReport As Worksheet
    Dim filterValue As String
    Dim lastRow As Long
    
    ' Thi?t l?p c�c d?i tu?ng Worksheet
    Set wsNewOrders = ThisWorkbook.Sheets("New Orders")
    Set wsReport = ThisWorkbook.Sheets("Report")
    
    ' L?y gi� tr? du?c ch?n t? drop-down menu ? � H10
    filterValue = wsNewOrders.Range("H10").Value
    
    ' X�c d?nh h�ng cu?i c�ng c?a c?t D trong sheet "New Orders"
    lastRow = wsNewOrders.Cells(wsNewOrders.Rows.Count, "D").End(xlUp).Row
    
     With wsNewOrders.Range("A3:D" & lastRow)
        .AutoFilter Field:=2, Criteria1:=filterValue
        ' Copy header
        .Resize(1, .Columns.Count).Copy wsReport.Cells(1, 1)
        ' Copy d? li?u (kh�ng bao g?m header)
        .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).SpecialCells(xlCellTypeVisible).Copy wsReport.Cells(wsReport.Rows.Count, 1).End(xlUp).Offset(1, 0)
    End With
    wsNewOrders.AutoFilterMode = False
    
End Sub

Sub Reset()
'Do NOT modify or delete this sub!
Sheets("Original Data").Columns("A:D").Copy Sheets("New Orders").Columns("A:D")
Sheets("Report").Cells.Clear
Sheets("Incomplete Orders").Cells.Clear
End Sub
