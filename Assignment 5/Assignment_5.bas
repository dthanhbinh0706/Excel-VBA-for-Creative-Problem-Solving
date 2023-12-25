Attribute VB_Name = "Module1"
Option Explicit

Sub FormatAndIncompleteOrders()
    Dim wsNewOrders As Worksheet
    Dim wsIncompleteOrders As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Ð?t tên cho các worksheet
    Set wsNewOrders = ThisWorkbook.Sheets("New Orders")
    Set wsIncompleteOrders = ThisWorkbook.Sheets("Incomplete Orders")

    ' Tìm hàng cu?i cùng có d? li?u trong c?t D c?a sheet "New Orders"
    lastRow = wsNewOrders.Cells(wsNewOrders.Rows.Count, "D").End(xlUp).Row

    ' 1. Xóa toàn b? format trên vùng d? li?u t? hàng 4 d?n hàng cu?i cùng c?a c?t A d?n D
    wsNewOrders.Range("A4:D" & lastRow).Style = "Normal"
    With wsNewOrders.Range("A4:A" & lastRow)
        .NumberFormat = "dd/mmm/yyyy"  ' Ð?nh d?ng ngày tháng theo ý mu?n
    End With

    ' 2. L?c và xóa các dòng có giá tr? tr?ng trong c?t B
    With wsNewOrders.Range("A3:D" & lastRow)
        .AutoFilter Field:=2, Criteria1:=""
        .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).SpecialCells(xlCellTypeVisible).ClearContents
    End With

    ' B? l?c
    wsNewOrders.AutoFilterMode = False
    
    ' Di chuy?n d? li?u lên n?u có hàng tr?ng
    For i = lastRow - 1 To 3 Step -1
        If IsEmpty(wsNewOrders.Cells(i, 2)) Then
            wsNewOrders.Rows(i).Delete
            wsNewOrders.Rows(i + 1).Copy wsNewOrders.Rows(i)
        End If
    Next i
    

    ' 3. L?c và sao chép các dòng có giá tr? tr?ng trong c?t C sang sheet "Incomplete Orders"
     With wsNewOrders.Range("A3:D" & lastRow)
        .AutoFilter Field:=3, Criteria1:="="
        ' Copy header
        .Resize(1, .Columns.Count).Copy wsIncompleteOrders.Cells(1, 1)
        ' Copy d? li?u (không bao g?m header)
        .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count).SpecialCells(xlCellTypeVisible).Copy wsIncompleteOrders.Cells(wsIncompleteOrders.Rows.Count, 1).End(xlUp).Offset(1, 0)
        ' Xóa các hàng d? li?u v?a du?c l?c ? sheet "New Orders"
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
    
    ' Thi?t l?p các d?i tu?ng Worksheet
    Set wsNewOrders = ThisWorkbook.Sheets("New Orders")
    Set wsReport = ThisWorkbook.Sheets("Report")
    
    ' L?y giá tr? du?c ch?n t? drop-down menu ? ô H10
    filterValue = wsNewOrders.Range("H10").Value
    
    ' Xác d?nh hàng cu?i cùng c?a c?t D trong sheet "New Orders"
    lastRow = wsNewOrders.Cells(wsNewOrders.Rows.Count, "D").End(xlUp).Row
    
     With wsNewOrders.Range("A3:D" & lastRow)
        .AutoFilter Field:=2, Criteria1:=filterValue
        ' Copy header
        .Resize(1, .Columns.Count).Copy wsReport.Cells(1, 1)
        ' Copy d? li?u (không bao g?m header)
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
