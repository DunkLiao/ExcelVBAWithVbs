Attribute VB_Name = "PivotWithPercentOfRow"
' ============================================================
' PivotWithPercentOfRow.bas
' 說明：使用 Excel VBA 自動建立顯示列佔比的樞紐分析表
' 功能：
'   1. 建立新活頁簿
'   2. 在「行銷預算」工作表填入行銷管道廣告費用示範資料
'   3. 建立樞紐分析表（列=行銷管道，欄=產品線）
'   4. 同時顯示廣告費用絕對值與每一列的佔比百分比
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 PivotWithPercentOfRow 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  As String = "行銷預算"
Const SHEET_PIVOT As String = "樞紐分析表"
Const PIVOT_NAME  As String = "列佔比樞紐"
Const OUTPUT_FILE As String = "20_PivotWithPercentOfRow.xlsx"

Sub PivotWithPercentOfRow()

    ' ── 範例資料（行銷管道、產品線、廣告費用）──────────────────
    ' 5 行銷管道 × 4 產品線 = 20 筆
    Dim arrChannels(19) As String
    Dim arrProducts(19) As String
    Dim arrBudgets(19)  As Long

    arrChannels(0)  = "社群媒體"   : arrProducts(0)  = "旗艦機" : arrBudgets(0)  = 320000
    arrChannels(1)  = "社群媒體"   : arrProducts(1)  = "標準機" : arrBudgets(1)  = 280000
    arrChannels(2)  = "社群媒體"   : arrProducts(2)  = "平板"   : arrBudgets(2)  = 175000
    arrChannels(3)  = "社群媒體"   : arrProducts(3)  = "穿戴"   : arrBudgets(3)  = 125000
    arrChannels(4)  = "搜尋廣告"   : arrProducts(4)  = "旗艦機" : arrBudgets(4)  = 250000
    arrChannels(5)  = "搜尋廣告"   : arrProducts(5)  = "標準機" : arrBudgets(5)  = 210000
    arrChannels(6)  = "搜尋廣告"   : arrProducts(6)  = "平板"   : arrBudgets(6)  = 140000
    arrChannels(7)  = "搜尋廣告"   : arrProducts(7)  = "穿戴"   : arrBudgets(7)  = 95000
    arrChannels(8)  = "電視廣告"   : arrProducts(8)  = "旗艦機" : arrBudgets(8)  = 480000
    arrChannels(9)  = "電視廣告"   : arrProducts(9)  = "標準機" : arrBudgets(9)  = 360000
    arrChannels(10) = "電視廣告"   : arrProducts(10) = "平板"   : arrBudgets(10) = 220000
    arrChannels(11) = "電視廣告"   : arrProducts(11) = "穿戴"   : arrBudgets(11) = 85000
    arrChannels(12) = "影音廣告"   : arrProducts(12) = "旗艦機" : arrBudgets(12) = 195000
    arrChannels(13) = "影音廣告"   : arrProducts(13) = "標準機" : arrBudgets(13) = 165000
    arrChannels(14) = "影音廣告"   : arrProducts(14) = "平板"   : arrBudgets(14) = 110000
    arrChannels(15) = "影音廣告"   : arrProducts(15) = "穿戴"   : arrBudgets(15) = 75000
    arrChannels(16) = "實體活動"   : arrProducts(16) = "旗艦機" : arrBudgets(16) = 150000
    arrChannels(17) = "實體活動"   : arrProducts(17) = "標準機" : arrBudgets(17) = 120000
    arrChannels(18) = "實體活動"   : arrProducts(18) = "平板"   : arrBudgets(18) = 85000
    arrChannels(19) = "實體活動"   : arrProducts(19) = "穿戴"   : arrBudgets(19) = 55000

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
    objDataSheet.Cells(1, 1).Value = "行銷管道"
    objDataSheet.Cells(1, 2).Value = "產品線"
    objDataSheet.Cells(1, 3).Value = "廣告費用"

    With objDataSheet.Range("A1:C1")
        .Font.Bold           = True
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 19
        objDataSheet.Cells(i + 2, 1).Value = arrChannels(i)
        objDataSheet.Cells(i + 2, 2).Value = arrProducts(i)
        objDataSheet.Cells(i + 2, 3).Value = arrBudgets(i)
    Next i

    objDataSheet.Columns("A:C").AutoFit

    ' ── 新增樞紐分析表工作表 ─────────────────────────────────────
    Set objPivotSheet  = objWorkbook.Sheets.Add()
    objPivotSheet.Name = SHEET_PIVOT
    objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

    ' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
    Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C21"))
    Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

    ' ── 設定列、欄欄位 ──────────────────────────────────────────
    Set objField = objPivot.PivotFields("行銷管道")
    objField.Orientation = xlRowField
    objField.Position    = 1

    Set objField = objPivot.PivotFields("產品線")
    objField.Orientation = xlColumnField
    objField.Position    = 1

    ' ── 第一個值欄位：廣告費用絕對金額 ─────────────────────────
    Set objField = objPivot.PivotFields("廣告費用")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "廣告費用（元）"

    ' ── 第二個值欄位：列佔比 ────────────────────────────────────
    Set objField = objPivot.PivotFields("廣告費用")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "列佔比（%）"

    Set objDataField = objPivot.DataFields("列佔比（%）")
    objDataField.Calculation  = xlPercentOfRow
    objDataField.NumberFormat = "0.0%"

    ' ── 加入說明標題 ─────────────────────────────────────────────
    objPivotSheet.Range("A1").Value = "列佔比樞紐分析表：各行銷管道在各產品線的廣告費用列佔比"
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
