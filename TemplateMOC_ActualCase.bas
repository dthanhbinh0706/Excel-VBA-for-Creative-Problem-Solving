Attribute VB_Name = "Module1"
Option Explicit
Private Sub Analyst_Click()
    '1) Khai bao bien
    Dim sourceSheet As Worksheet
    ' La sheet "FormatedData"
    Dim targetSheet As Worksheet
    ' La sheet "MgmtFeeStore"
    Dim wsMgmtFeeStore As Worksheet
    
    '--------------------------------------------------------------------------------------------------------------
    '2) Thiet lap: Sheet "MgmtFeeStore"
    Set wsMgmtFeeStore = ThisWorkbook.Sheets("MgmtFeeStore")
    
    '--------------------------------------------------------------------------------------------------------------
    '3) Thiet lap: Sheet "FormatedData"
    '   3.1) Tìm sheet (Receipt-report) tu nguon coi da co chua
    On Error Resume Next
    Set sourceSheet = Worksheets("Receipt-report")
    On Error GoTo 0
    '   Neu sheet (Receipt-report) tu nguon khong tim thay, thông báo và thoát
    If sourceSheet Is Nothing Then
        MsgBox "Sheet 'Receipt-report' did not found in your folder", vbExclamation
        Exit Sub
    End If
    '   3.2) Sau khi tìm ra sheet (Receipt-report) tu nguon, bat dau copy du lieu
    '   Tao sheet moi (FormatedData) hoac su dung sheet da ton tai
    On Error Resume Next
    Set targetSheet = Worksheets("FormatedData")
    On Error GoTo 0
    '   Neu sheet không ton tai, tao moi
    If targetSheet Is Nothing Then
        Set targetSheet = Sheets.Add(After:=Sheets(Sheets.Count))
        targetSheet.Name = "FormatedData"
    Else
    '   Neu sheet dã ton tai, xóa du lieu cu
        targetSheet.Cells.Clear
    End If
    '   3.3) Copy du lieu tu sheet nguon (Receipt-report) sang sheet dích ("FormatedData")
    sourceSheet.UsedRange.Copy targetSheet.Range("A1")
    
    '--------------------------------------------------------------------------------------------------------------
    Application.ScreenUpdating = False
    '4) Xu ly cac cot theo yeu cau: Sheet "FormatedData"
    '   4.1) Cot K ("Remark")
    '       + Goi hàm RemoveInviOrBook() de loai bo cac chuoi co trong ds
    RemoveInviOrBook targetSheet
    '       + Goi hàm ReplaceCRMMKT() và ReplacePMHCRM() de thay the cac chuoi thanh CRMMKT hoac PMHCRM
    ReplaceCRMMKT targetSheet
    ReplacePMHCRM targetSheet
    '   4.2) Them Cot L ("Denomination") theo cac yeu cau
    AddDenominationColumn targetSheet
    '   4.3) Them Cot M ("MgmtFeeStore") theo cac yeu cau
    AddMgmtFeeStoreColumn targetSheet, wsMgmtFeeStore
    AddMgmtFeeStoreWithConditions targetSheet, wsMgmtFeeStore
    '   4.4) Them Cot N ("ServiceFee") theo cac yeu cau
    AddServiceFeeColumn targetSheet
    '   4.5) Them Cot O ("VAT") theo cac yeu cau
    AddVATColumn targetSheet
    '   4.6) Them Cot P ("TotalServiceFee") theo cac yeu cau
    AddTotalServiceFeeColumn targetSheet
    '   4.7) Them Cot Q ("TotalServiceFee") theo cac yeu cau
    AddTotalAfterFeeColumn targetSheet
    
    Application.ScreenUpdating = True
    '--------------------------------------------------------------------------------------------------------------
    '5) Tao PivotTable tai sheet moi "Pivot" dua tren data sheet "FormatedData"
    CreatePivotTableForSheet targetSheet
    
    '--------------------------------------------------------------------------------------------------------------
    '6) Hien thong thông báo khi hoàn thành
    MsgBox "The analyst process has been successfully completed, please select OK to finish!", vbInformation
End Sub
Sub RemoveInviOrBook(ws As Worksheet)
    Dim rng As Range
    Dim cell As Range
    Dim replaceList As Variant
    Dim replaceString As Variant
    
    ' Danh sách chu?i c?n thay th?
    replaceList = Array("2190 Book ", "2500 Book ", "970 Book ", "256 Individual ", "610 Individual ", "202 Individual ")
    ' Xác d?nh ph?m vi c?n thay th? trong c?t K
    Set rng = ws.Range("K2:K" & ws.Cells(ws.Rows.Count, "K").End(xlUp).Row)
    
    ' Duy?t qua t?ng ô trong ph?m vi
    For Each cell In rng
        ' Duy?t qua t?ng chu?i c?n thay th?
        For Each replaceString In replaceList
            ' Ki?m tra xem ô có ch?a chu?i c?n thay th? không
            If InStr(1, cell.Value, replaceString) > 0 Then
                ' Thay th? chu?i b?ng chu?i tr?ng
                cell.Value = Replace(cell.Value, replaceString, "")
                Exit For ' Thoát kh?i vòng l?p khi dã th?c hi?n thay th?
            End If
        Next replaceString
    Next cell
End Sub

Sub ReplaceCRMMKT(ws As Worksheet)
    Dim rng As Range
    Dim cell As Range
    Dim regex As Object
    
    ' Su dung bieu thuc chính quy de tìm kiem và thay the
    Set regex = CreateObject("VBScript.RegExp")
    ' Dat reges o che do tiem kiem toan bo chuoi (khong chi tim kiem lan dau tien)
    regex.Global = True
    regex.Pattern = "CRMMKT\d*"
    
    ' Xác dinh pham vi can thay the trong cot K
    Set rng = ws.Range("K2:K" & ws.Cells(ws.Rows.Count, "K").End(xlUp).Row)
    
    ' Duyet qua tung o trong pham vi
    For Each cell In rng
        'Su dung Test de kiem tra xem gia tri trong o co khop voi biet thuc chinh quy hay khong
        If regex.Test(cell.Value) Then
            ' Thuc hien thay the neu co su khop
            cell.Value = regex.Replace(cell.Value, "CRMMKT")
        End If
    Next cell
    
End Sub
Sub ReplacePMHCRM(ws As Worksheet)
    Dim rng As Range
    Dim cell As Range
    Dim regex As Object
    
    ' Su dung bieu thuc chính quy de tìm kiem và thay the
    Set regex = CreateObject("VBScript.RegExp")
    ' Dat reges o che do tiem kiem toan bo chuoi (khong chi tim kiem lan dau tien)
    regex.Global = True
    regex.Pattern = "PMHCRM\d*"
    
    ' Xác dinh pham vi can thay the trong cot K
    Set rng = ws.Range("K2:K" & ws.Cells(ws.Rows.Count, "K").End(xlUp).Row)
    
    ' Duyet qua tung o trong pham vi
    For Each cell In rng
        'Su dung Test de kiem tra xem gia tri trong o co khop voi biet thuc chinh quy hay khong
        If regex.Test(cell.Value) Then
            ' Thuc hien thay the neu co su khop
            cell.Value = regex.Replace(cell.Value, "PMHCRM")
        End If
    Next cell
    
End Sub
Sub AddDenominationColumn(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim originalValue As String
    Dim result As String
    
    ' Tìm dòng cuoi cùng có du lieu trong cot K caa worksheet
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    ' Tao moi cot "Denomination" bên phai cot K
    ws.Columns("L:L").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ' Ðat tên cho cot là "Denomination"
    ws.Cells(1, 12).Value = "Denomination"
    
    ' Duyet qua tung o trong pham vi
    For i = 2 To lastRow
        ' Lay giá tri tung ô trong pham vi
        originalValue = ws.Cells(i, 11).Value
        ' Thuc hien cong thuc tach gia tri: 50 trong "Crescent Mall Gift Voucher - 50.000 VND"
        result = Mid(originalValue, InStr(originalValue, "- ") + 2, InStr(Mid(originalValue, InStr(originalValue, "- ") + 2), ".") - 1)
        ' Kiem tra ket qua và thêm "000000" neu la so 1 hoac "000" cho cac gia tri con lai
        If result = "1" Then
            result = result & "000000"
        Else
            result = result & "000"
        End If
        ' Gán giá tri vào cot moi tao
        ws.Cells(i, 12).Value = result
    Next i

End Sub
Sub AddMgmtFeeStoreColumn(wsFormattedData As Worksheet, wsMgmtFeeStore As Worksheet)
    Dim lastRowFormattedData As Long
    Dim lastRowMgmtFeeStore As Long
    Dim i As Long

    ' Tìm dong cuoi cung co du lieu trong moi bang
    lastRowFormattedData = wsFormattedData.Cells(wsFormattedData.Rows.Count, "A").End(xlUp).Row
    lastRowMgmtFeeStore = wsMgmtFeeStore.Cells(wsMgmtFeeStore.Rows.Count, "A").End(xlUp).Row
    
    ' Tao moi cot "MgmtFeeStore" bên phai cot L
    wsFormattedData.Columns("M:M").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ' Ðat tên cho cot là "MgmtFeeStore"
    wsFormattedData.Cells(1, 13).Value = "MgmtFeeStore"
    
    ' Duyet qua tung o trong pham vi
    For i = 2 To lastRowFormattedData
        ' Lay giá tri ConsumedStore[i]
        Dim consumedStore As String
        consumedStore = wsFormattedData.Cells(i, "H").Value

        ' Tìm kiem trong bang MgmtFeeStore
        Dim mgmtFee As Variant
        mgmtFee = Application.VLookup(consumedStore, wsMgmtFeeStore.Range("A:B"), 2, False)

        ' Kiem tra ResignDate[i] có trong hay không
        If wsMgmtFeeStore.Cells(Application.Match(consumedStore, wsMgmtFeeStore.Columns(1), 0), "C").Value = "" Then
            ' Neu trong, dien giá tri ManagementFee tuong ung
            wsFormattedData.Cells(i, wsFormattedData.Cells(1, wsFormattedData.Columns.Count).End(xlToLeft).Column).Value = mgmtFee
        Else
            ' Neu không trong, dien giá tri ""
            wsFormattedData.Cells(i, wsFormattedData.Cells(1, wsFormattedData.Columns.Count).End(xlToLeft).Column).Value = ""
        End If
    Next i
    
End Sub
Sub AddMgmtFeeStoreWithConditions(wsFormattedData As Worksheet, wsMgmtFeeStore As Worksheet)
    ' 1. Khai bao bien
    Dim lastRowFormattedData As Long
    Dim lastRowMgmtFeeStore As Long
    Dim i As Long
    
    
    ' 2. Tìm dong cuoi cung co du lieu trong moi bang
    lastRowFormattedData = wsFormattedData.Cells(wsFormattedData.Rows.Count, "A").End(xlUp).Row
    lastRowMgmtFeeStore = wsMgmtFeeStore.Cells(wsMgmtFeeStore.Rows.Count, "A").End(xlUp).Row
    
    ' 3. Bat dau vòng lap qua tung dòng trong bang wsFormattedData
    For i = 2 To lastRowFormattedData
        
        ' 4. Lay giá tri MgmtFeeStore[i] trong bang wsFormattedData
        Dim mgmtFeeStore As Variant
        mgmtFeeStore = wsFormattedData.Cells(i, "M").Value
        
        ' 5. Su dung vòng lap Do While de kiem tra xem mgmtFeeStore có trong hay không
        ' Neu mgmtFeeStore không rong, vòng lap se không duoc thuc hien
        Do While mgmtFeeStore = ""
            ' 6. Neu mgmtFeeStore trong
            ' Lay giá tri ConsumedStore[i] và ConsumedAt[i]
            Dim consumedStore As String
            Dim consumedAt As Date
            consumedStore = wsFormattedData.Cells(i, "H").Value
            ' Chuyen doi mot giá tri sang kieu du lieu Date
            consumedAt = CDate(wsFormattedData.Cells(i, "G").Value)
            
            ' Tìm kiem trong bang MgmtFeeStore
            Dim mgmtFee As Variant
            Dim resignDate As Date
            Dim endExtend As Date
            mgmtFee = Application.VLookup(consumedStore, wsMgmtFeeStore.Range("A:B"), 2, False)
            resignDate = CDate(wsMgmtFeeStore.Cells(Application.Match(consumedStore, wsMgmtFeeStore.Columns(1), 0), "C").Value)
            endExtend = CDate(wsMgmtFeeStore.Cells(Application.Match(consumedStore, wsMgmtFeeStore.Columns(1), 0), "D").Value)
            
            ' Kiem tra dieu kien ResignDate[i] <= ConsumedAt[i] <= EndExtend[i]
            If resignDate <= consumedAt And consumedAt <= endExtend Then
                ' 7. Neu True
                ' Gán giá tri mgmtFee vào cot "M" o dòng thu i trong bang wsFormattedData và thoát khoi vong lap Do While
                wsFormattedData.Cells(i, "M").Value = mgmtFee
                Exit Do
            Else
                ' 9. Neu False
                ' Gán giá tri 0 vào cot "M" o dòng thu i trong bang wsFormattedData
                wsFormattedData.Cells(i, "M").Value = 0
                mgmtFeeStore = wsFormattedData.Cells(i, "M").Value
            End If
        Loop
    Next i
End Sub
Sub AddServiceFeeColumn(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim originalValue As Variant
    Dim result As Variant

    ' Tìm dòng cuoi cùng có du lieu trong cot M cua worksheet
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
    ' Tao moi cot "MgmtFee" bên phai cot L
    ws.Columns("N:N").Insert Shift:=xlToRight
    ' Ðat tên cho cot moi là "MgmtFee"
    ws.Cells(1, 14).Value = "ServiceFee"
    
    ' Duyet qua tung o trong pham vi
    For i = 2 To lastRow
        ' Lay giá tri tung ô trong pham vi
        originalValue = ws.Cells(i, 13).Value
        ' Thuc hien công thuc de tao giá tri cho "MgmtFeeStore"
        result = originalValue * ws.Cells(i, 12).Value
        ' Gán giá tri vào cot "ManagementFee"
        ws.Cells(i, 14).Value = result
    Next i
End Sub
Sub AddVATColumn(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim originalValue As Variant
    Dim result As Variant

    ' Tìm dòng cuoi cùng có du lieu trong cot N cua worksheet
    lastRow = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
    ' Tao moi cot "MgmtFee" bên phai cot N
    ws.Columns("O:O").Insert Shift:=xlToRight
    ' Ðat tên cho cot moi là "VAT"
    ws.Cells(1, 15).Value = "VAT"
    
    ' Duyet qua tung vong trong pham vi
    For i = 2 To lastRow
        ' Lay gia tri tung o trong pham vi
        originalValue = ws.Cells(i, 14).Value
        ' Thuc hien cong thuc de tao gia tri cho "VAT"
        result = originalValue * 0.1
        ' Gán giá tri vào cot "VAT"
        ws.Cells(i, 15).Value = result
    Next i
End Sub
Sub AddTotalServiceFeeColumn(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim originalValue As Variant
    Dim result As Variant

    ' Tìm dòng cuoi cùng có du lieu trong cot O cua worksheet
    lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
    ' Tao moi cot "TotalServiceFee" bên phai cot O
    ws.Columns("P:P").Insert Shift:=xlToRight
    ' Ðat tên cho cot moi là "TotalServiceFee"
    ws.Cells(1, 16).Value = "TotalServiceFee"
    
    ' Duyet qua tung dòng trong pham vi
    For i = 2 To lastRow
        ' Lay giá tri tung ô trong pham vi
        originalValue = ws.Cells(i, 15).Value
        ' Thuc hien công thuc de tao giá tri cho "TotalServiceFee"
        result = originalValue + ws.Cells(i, 14).Value
        ' Gán giá tri vào cot "TotalServiceFee"
        ws.Cells(i, 16).Value = result
    Next i
End Sub
Sub AddTotalAfterFeeColumn(ws As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim originalValue As Variant
    Dim result As Variant

    ' Tìm dòng cuoi cùng có du lieu trong cot P cua worksheet
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    ' Tao moi cot "TotalAfterFee" bên phai cot P
    ws.Columns("Q:Q").Insert Shift:=xlToRight
    ' Ðat tên cho cot moi là "TotalAfterFee"
    ws.Cells(1, 17).Value = "TotalAfterFee"
    
    ' Duyet qua tung dòng trong pham vi
    For i = 2 To lastRow
        ' Lay giá tri tung ô trong pham vi
        originalValue = ws.Cells(i, 12).Value
        ' Thuc hien công thuc de tao giá tri cho "TotalAfterFee"
        result = originalValue - ws.Cells(i, 16).Value
        ' Gán giá tri vao cot "TotalAfterFee"
        ws.Cells(i, 17).Value = result
    Next i
End Sub
Sub CreatePivotTableForSheet(targetSheet As Worksheet)
    Dim sourceRange As Range
    Dim pivotCache As pivotCache
    Dim pivotTable As pivotTable
    Dim pivotSheet As Worksheet
    Dim existingPivotTable As pivotTable

    ' UsedRange khong bao gom cac o trô'ng o cuôi sheet hoac cac dòng cot không su dung.
    Set sourceRange = targetSheet.UsedRange

    ' Kiem tra xem PivotSheet có ton tai chua, neu chua thì tao moi
    On Error Resume Next
    Set pivotSheet = Sheets("Pivot")
    On Error GoTo 0
    If pivotSheet Is Nothing Then
        Set pivotSheet = Sheets.Add(After:=Sheets(Sheets.Count))
        pivotSheet.Name = "Pivot"
    End If
    ' Kiem tra xem PivotTable1 dã ton tai chua, neu có thì xóa
    On Error Resume Next
    Set existingPivotTable = pivotSheet.PivotTables("PivotTable1")
    On Error GoTo 0

    If Not existingPivotTable Is Nothing Then
        existingPivotTable.TableRange2.Clear
        existingPivotTable.PivotTableWizard TableDestination:=pivotSheet.Cells(3, 1)
        existingPivotTable.Name = "PivotTable1" ' Ðoi tên thành "PivotTable1"
    Else
        ' Tao PivotCache và PivotTable neu chua ton tai
        Set pivotCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=sourceRange, Version:=xlPivotTableVersion15)
        Set pivotTable = pivotCache.CreatePivotTable(TableDestination:=pivotSheet.Cells(3, 1), TableName:="PivotTable1", DefaultVersion:=xlPivotTableVersion15)

        ' Cau hình PivotTable
        With pivotTable
            .ColumnGrand = False
            ' ... (các cau hình khác cua PivotTable)
        End With

        ' Thêm truong "ConsumedStore" vào PivotTable
        With pivotTable.PivotFields("ConsumedStore")
            .Orientation = xlRowField
            .Position = 1
        End With

        ' Thêm truong "MgmtFeeStore" vào PivotTable
        With pivotTable.PivotFields("MgmtFeeStore")
            .Orientation = xlRowField
            .NumberFormat = "0%"
        End With
        
        ' Thêm truong "Remark" vào PivotTable
        With pivotTable.PivotFields("Remark")
            .Orientation = xlRowField
        End With

        ' Ðem so luong Remark
        pivotTable.AddDataField pivotTable.PivotFields("ReceiptUUID"), "Count of ReceiptUUID", xlCount
        ' T?ng Denomination
        pivotTable.AddDataField pivotTable.PivotFields("Denomination"), "Sum of Denomination", xlSum
        pivotTable.PivotFields("Sum of Denomination").NumberFormat = "#,##0"
        ' T?ng ServiceFee
        pivotTable.AddDataField pivotTable.PivotFields("ServiceFee"), "Sum of ServiceFee", xlSum
        pivotTable.PivotFields("Sum of ServiceFee").NumberFormat = "#,##0"
        ' T?ng VAT
        pivotTable.AddDataField pivotTable.PivotFields("VAT"), "Sum of VAT", xlSum
        pivotTable.PivotFields("Sum of VAT").NumberFormat = "#,##0"
        ' T?ng TotalServiceFee
        pivotTable.AddDataField pivotTable.PivotFields("TotalServiceFee"), "Sum of TotalServiceFee", xlSum
        pivotTable.PivotFields("Sum of TotalServiceFee").NumberFormat = "#,##0"
        ' T?ng TotalAfterFee
        pivotTable.AddDataField pivotTable.PivotFields("TotalAfterFee"), "Sum of TotalAfterFee", xlSum
        pivotTable.PivotFields("Sum of TotalAfterFee").NumberFormat = "#,##0"
        
    End If
End Sub

Private Sub ImportFile_Click()
Dim filePath As String
    Dim importSheet As Worksheet
    Dim ws As Worksheet

    ' Chon file Excel de import
    filePath = Application.GetOpenFilename("Excel Files (*.xlsx), *.xlsx")

    ' Kiem tra nguoi dung da chen file chua
    If filePath <> "False" Then
        ' Tao sheet moi có tên "Receipt-report"
        Set importSheet = Sheets.Add(After:=Sheets(Sheets.Count))
        importSheet.Name = "Receipt-report"

        ' Sao chép du lieu tu file dã chen vào sheet moi
        Set ws = ThisWorkbook.Sheets("Receipt-report")
        Workbooks.Open (filePath)
        ActiveWorkbook.Sheets(1).UsedRange.Copy ws.Range("A1")
        ActiveWorkbook.Close SaveChanges:=False

        ' Hien thi thông báo thành công
        MsgBox "Import File Successfully!", vbInformation
    Else
        ' Hien thi thông báo neu nguoi dùng không chen file
        MsgBox "The Import Process was cancelled", vbExclamation
    End If
End Sub

Private Sub Reset_Click()
    On Error Resume Next
    Application.DisplayAlerts = False ' Tat thông báo xác nhan xóa sheet

    If Not SheetExists("Receipt-report") Then
        If MsgBox("The 'Receipt-report' sheet does not exist. Continue to next sheet?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub ' Neu nguoi dùng chon No, thoát khoi hàm
        End If
    Else
        Sheets("Receipt-report").Delete
    End If

    If Not SheetExists("FormatedData") Then
        If MsgBox("The 'FormatedData' sheet does not exist. Continue to next sheet?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub ' N?u ngu?i dùng ch?n No, thoát kh?i hàm
        End If
    Else
        Sheets("FormatedData").Delete
    End If

    If Not SheetExists("Pivot") Then
        If MsgBox("The 'Pivot' sheet does not exist. Continue to next sheet?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub ' N?u ngu?i dùng ch?n No, thoát kh?i hàm
        End If
    Else
        Sheets("Pivot").Delete
    End If

    Application.DisplayAlerts = True ' B?t l?i thông báo xác nh?n xóa sheet
    On Error GoTo 0
End Sub

Function SheetExists(sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function


