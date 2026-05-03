Attribute VB_Name = "CreateDoughnutChart"
' ============================================================
' CreateDoughnutChart.bas
' 說明：使用 Excel VBA 自動建立環圈圖範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（年度預算分配）
'   3. 插入環圈圖
'   4. 設定圖表標題、百分比資料標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreateDoughnutChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE As String = "2025 年度預算分配"
Const SHEET_NAME  As String = "預算分配"
Const OUTPUT_FILE As String = "DoughnutChartExample.xlsx"

Sub CreateDoughnutChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    Dim arrItems(5) As String
    arrItems(0) = "研發費用" : arrItems(1) = "行銷費用" : arrItems(2) = "人事費用"
    arrItems(3) = "行政費用" : arrItems(4) = "設備費用" : arrItems(5) = "其他費用"

    Dim arrBudget(5) As Long
    arrBudget(0) = 30 : arrBudget(1) = 20 : arrBudget(2) = 25
    arrBudget(3) = 10 : arrBudget(4) = 10 : arrBudget(5) = 5

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook As Workbook
    Dim objSheet    As Worksheet
    Dim objChartObj As ChartObject
    Dim objChart    As Chart
    Dim savePath    As String
    Dim i           As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook = Workbooks.Add()
    Set objSheet    = objWorkbook.Sheets(1)
    objSheet.Name   = SHEET_NAME

    ' ── 寫入標題列 ──────────────────────────────────────────────
    objSheet.Cells(1, 1).Value = "費用項目"
    objSheet.Cells(1, 2).Value = "佔比（%）"

    With objSheet.Range("A1:B1")
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 5
        objSheet.Cells(i + 2, 1).Value = arrItems(i)
        objSheet.Cells(i + 2, 2).Value = arrBudget(i)
    Next i

    objSheet.Columns("A:B").AutoFit

    ' ── 插入環圈圖 ───────────────────────────────────────────────
    Set objChartObj = objSheet.ChartObjects.Add(200, 20, 400, 300)
    Set objChart    = objChartObj.Chart

    objChart.ChartType = xlDoughnut
    objChart.SetSourceData objSheet.Range("A1:B7")

    ' ── 圖表格式設定 ────────────────────────────────────────────
    objChart.HasTitle        = True
    objChart.ChartTitle.Text = CHART_TITLE
    objChart.ChartTitle.Font.Size = 14
    objChart.ChartTitle.Font.Bold = True

    ' 顯示百分比資料標籤 + 類別名稱
    With objChart.SeriesCollection(1)
        .HasDataLabels = True
        With .DataLabels
            .ShowPercentage   = True
            .ShowCategoryName = True
            .ShowValue        = False
        End With
    End With

    objChart.HasLegend = True

    ' ── 儲存 ────────────────────────────────────────────────────
    objWorkbook.SaveAs savePath, xlOpenXMLWorkbook

    Set objChart    = Nothing
    Set objChartObj = Nothing
    Set objSheet    = Nothing
    Set objWorkbook = Nothing

    MsgBox "完成！檔案已儲存至：" & savePath

End Sub
