Attribute VB_Name = "PivotWithRunningTotal"
' ============================================================
' PivotWithRunningTotal.bas
' 說明：使用 Excel VBA 自動建立顯示累計加總的樞紐分析表
' 功能：
'   1. 建立新活頁簿
'   2. 在「門市業績」工作表填入月份門市銷售示範資料
'   3. 建立樞紐分析表（列=月份，欄=門市）
'   4. 同時顯示當月銷售額及月份維度的累計銷售額
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 PivotWithRunningTotal 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  As String = "門市業績"
Const SHEET_PIVOT As String = "樞紐分析表"
Const PIVOT_NAME  As String = "累計加總樞紐"
Const OUTPUT_FILE As String = "12_PivotWithRunningTotal.xlsx"

Sub PivotWithRunningTotal()

    ' ── 範例資料（月份、門市、銷售額）──────────────────────────
    ' 12 月 × 2 門市 = 24 筆
    Dim arrMonths(23) As Long
    Dim arrStores(23) As String
    Dim arrAmounts(23) As Long

    arrMonths(0)  = 1  : arrStores(0)  = "信義店" : arrAmounts(0)  = 85000
    arrMonths(1)  = 1  : arrStores(1)  = "西門店" : arrAmounts(1)  = 72000
    arrMonths(2)  = 2  : arrStores(2)  = "信義店" : arrAmounts(2)  = 78000
    arrMonths(3)  = 2  : arrStores(3)  = "西門店" : arrAmounts(3)  = 65000
    arrMonths(4)  = 3  : arrStores(4)  = "信義店" : arrAmounts(4)  = 95000
    arrMonths(5)  = 3  : arrStores(5)  = "西門店" : arrAmounts(5)  = 81000
    arrMonths(6)  = 4  : arrStores(6)  = "信義店" : arrAmounts(6)  = 102000
    arrMonths(7)  = 4  : arrStores(7)  = "西門店" : arrAmounts(7)  = 88000
    arrMonths(8)  = 5  : arrStores(8)  = "信義店" : arrAmounts(8)  = 118000
    arrMonths(9)  = 5  : arrStores(9)  = "西門店" : arrAmounts(9)  = 97000
    arrMonths(10) = 6  : arrStores(10) = "信義店" : arrAmounts(10) = 125000
    arrMonths(11) = 6  : arrStores(11) = "西門店" : arrAmounts(11) = 108000
    arrMonths(12) = 7  : arrStores(12) = "信義店" : arrAmounts(12) = 132000
    arrMonths(13) = 7  : arrStores(13) = "西門店" : arrAmounts(13) = 115000
    arrMonths(14) = 8  : arrStores(14) = "信義店" : arrAmounts(14) = 128000
    arrMonths(15) = 8  : arrStores(15) = "西門店" : arrAmounts(15) = 110000
    arrMonths(16) = 9  : arrStores(16) = "信義店" : arrAmounts(16) = 145000
    arrMonths(17) = 9  : arrStores(17) = "西門店" : arrAmounts(17) = 124000
    arrMonths(18) = 10 : arrStores(18) = "信義店" : arrAmounts(18) = 158000
    arrMonths(19) = 10 : arrStores(19) = "西門店" : arrAmounts(19) = 135000
    arrMonths(20) = 11 : arrStores(20) = "信義店" : arrAmounts(20) = 172000
    arrMonths(21) = 11 : arrStores(21) = "西門店" : arrAmounts(21) = 148000
    arrMonths(22) = 12 : arrStores(22) = "信義店" : arrAmounts(22) = 198000
    arrMonths(23) = 12 : arrStores(23) = "西門店" : arrAmounts(23) = 168000

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook   As Workbook
    Dim objDataSheet  As Worksheet
    Dim objPivotSheet As Worksheet
    Dim objCache      As PivotCache
    Dim objPivot      As PivotTable
    Dim objField      As PivotField
    Dim objDataField  As PivotField
    Dim savePath      As String
    Dim i             As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook   = Workbooks.Add()
    Set objDataSheet  = objWorkbook.Sheets(1)
    objDataSheet.Name = SHEET_DATA

    ' ── 寫入標題列 ──────────────────────────────────────────────
    objDataSheet.Cells(1, 1).Value = "月份"
    objDataSheet.Cells(1, 2).Value = "門市"
    objDataSheet.Cells(1, 3).Value = "銷售額"

    With objDataSheet.Range("A1:C1")
        .Font.Bold           = True
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 23
        objDataSheet.Cells(i + 2, 1).Value = arrMonths(i)
        objDataSheet.Cells(i + 2, 2).Value = arrStores(i)
        objDataSheet.Cells(i + 2, 3).Value = arrAmounts(i)
    Next i

    objDataSheet.Columns("A:C").AutoFit

    ' ── 新增樞紐分析表工作表 ─────────────────────────────────────
    Set objPivotSheet  = objWorkbook.Sheets.Add()
    objPivotSheet.Name = SHEET_PIVOT
    objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

    ' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
    Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C25"))
    Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

    ' ── 設定列、欄欄位 ──────────────────────────────────────────
    Set objField = objPivot.PivotFields("月份")
    objField.Orientation = xlRowField
    objField.Position    = 1

    Set objField = objPivot.PivotFields("門市")
    objField.Orientation = xlColumnField
    objField.Position    = 1

    ' ── 第一個值欄位：當月銷售額 ────────────────────────────────
    Set objField = objPivot.PivotFields("銷售額")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "當月銷售額"

    ' ── 第二個值欄位：累計銷售額 ────────────────────────────────
    Set objField = objPivot.PivotFields("銷售額")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "累計銷售額"

    ' ── 設定累計加總計算方式（以月份為基準欄位）────────────────
    Set objDataField = objPivot.DataFields("累計銷售額")
    objDataField.Calculation = xlRunningTotal
    objDataField.BaseField   = "月份"

    ' ── 加入說明標題 ─────────────────────────────────────────────
    objPivotSheet.Range("A1").Value = "累計加總樞紐分析表：各門市當月銷售額及月份累計加總"
    With objPivotSheet.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
    End With

    ' ── 儲存 ────────────────────────────────────────────────────
    objWorkbook.SaveAs savePath, xlOpenXMLWorkbook

    Set objDataField  = Nothing
    Set objField      = Nothing
    Set objPivot      = Nothing
    Set objCache      = Nothing
    Set objPivotSheet = Nothing
    Set objDataSheet  = Nothing
    Set objWorkbook   = Nothing

    MsgBox "完成！檔案已儲存至：" & savePath

End Sub
