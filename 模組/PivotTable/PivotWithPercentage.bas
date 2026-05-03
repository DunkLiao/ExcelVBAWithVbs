Attribute VB_Name = "PivotWithPercentage"
' ============================================================
' PivotWithPercentage.bas
' 說明：使用 Excel VBA 自動建立以百分比總計顯示的樞紐分析表
' 功能：
'   1. 建立新活頁簿
'   2. 在「市佔資料」工作表填入品牌通路銷售示範資料
'   3. 建立樞紐分析表，值欄位以「佔總計百分比」方式顯示
'   4. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 PivotWithPercentage 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  As String = "市佔資料"
Const SHEET_PIVOT As String = "樞紐分析表"
Const PIVOT_NAME  As String = "百分比樞紐"
Const OUTPUT_FILE As String = "06_PivotWithPercentage.xlsx"

Sub PivotWithPercentage()

    ' ── 範例資料（品牌、通路、銷售額）──────────────────────────
    Dim arrBrands(15)   As String
    Dim arrChannels(15) As String
    Dim arrAmounts(15)  As Long

    arrBrands(0)  = "品牌A" : arrChannels(0)  = "直營門市" : arrAmounts(0)  = 320000
    arrBrands(1)  = "品牌A" : arrChannels(1)  = "電商平台" : arrAmounts(1)  = 285000
    arrBrands(2)  = "品牌A" : arrChannels(2)  = "代理通路" : arrAmounts(2)  = 178000
    arrBrands(3)  = "品牌A" : arrChannels(3)  = "展會直銷" : arrAmounts(3)  = 95000
    arrBrands(4)  = "品牌B" : arrChannels(4)  = "直營門市" : arrAmounts(4)  = 210000
    arrBrands(5)  = "品牌B" : arrChannels(5)  = "電商平台" : arrAmounts(5)  = 365000
    arrBrands(6)  = "品牌B" : arrChannels(6)  = "代理通路" : arrAmounts(6)  = 142000
    arrBrands(7)  = "品牌B" : arrChannels(7)  = "展會直銷" : arrAmounts(7)  = 68000
    arrBrands(8)  = "品牌C" : arrChannels(8)  = "直營門市" : arrAmounts(8)  = 152000
    arrBrands(9)  = "品牌C" : arrChannels(9)  = "電商平台" : arrAmounts(9)  = 198000
    arrBrands(10) = "品牌C" : arrChannels(10) = "代理通路" : arrAmounts(10) = 89000
    arrBrands(11) = "品牌C" : arrChannels(11) = "展會直銷" : arrAmounts(11) = 43000
    arrBrands(12) = "品牌D" : arrChannels(12) = "直營門市" : arrAmounts(12) = 185000
    arrBrands(13) = "品牌D" : arrChannels(13) = "電商平台" : arrAmounts(13) = 142000
    arrBrands(14) = "品牌D" : arrChannels(14) = "代理通路" : arrAmounts(14) = 115000
    arrBrands(15) = "品牌D" : arrChannels(15) = "展會直銷" : arrAmounts(15) = 57000

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
    objDataSheet.Cells(1, 1).Value = "品牌"
    objDataSheet.Cells(1, 2).Value = "通路"
    objDataSheet.Cells(1, 3).Value = "銷售額"

    With objDataSheet.Range("A1:C1")
        .Font.Bold           = True
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 15
        objDataSheet.Cells(i + 2, 1).Value = arrBrands(i)
        objDataSheet.Cells(i + 2, 2).Value = arrChannels(i)
        objDataSheet.Cells(i + 2, 3).Value = arrAmounts(i)
    Next i

    objDataSheet.Columns("A:C").AutoFit

    ' ── 新增樞紐分析表工作表 ─────────────────────────────────────
    Set objPivotSheet  = objWorkbook.Sheets.Add()
    objPivotSheet.Name = SHEET_PIVOT
    objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

    ' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
    Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C17"))
    Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

    ' ── 設定列、欄欄位 ──────────────────────────────────────────
    Set objField = objPivot.PivotFields("品牌")
    objField.Orientation = xlRowField
    objField.Position    = 1

    Set objField = objPivot.PivotFields("通路")
    objField.Orientation = xlColumnField
    objField.Position    = 1

    ' ── 設定值欄位，並將顯示方式設為「佔總計百分比」──────────────
    Set objField = objPivot.PivotFields("銷售額")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "市佔率 (%)"

    Set objDataField = objPivot.DataFields("市佔率 (%)")
    objDataField.Calculation  = xlPercentOfTotal
    objDataField.NumberFormat = "0.00%"

    ' ── 加入說明標題 ─────────────────────────────────────────────
    objPivotSheet.Range("A1").Value = "百分比樞紐分析表：各品牌各通路銷售額佔總計的比率"
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
