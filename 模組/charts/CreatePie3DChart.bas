Attribute VB_Name = "CreatePie3DChart"
' ============================================================
' CreatePie3DChart.bas
' 說明：使用 Excel VBA 自動建立 3D 圓餅圖範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（公司資源分配比例）
'   3. 插入 3D 圓餅圖
'   4. 設定圖表標題、百分比資料標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreatePie3DChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE As String = "2025 年公司資源分配（3D）"
Const SHEET_NAME  As String = "資源分配"
Const OUTPUT_FILE As String = "Pie3DChartExample.xlsx"

Sub CreatePie3DChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    Dim arrItems(4) As String
    arrItems(0) = "技術研發" : arrItems(1) = "市場推廣" : arrItems(2) = "人才培育"
    arrItems(3) = "基礎設施" : arrItems(4) = "客戶服務"

    Dim arrPercent(4) As Long
    arrPercent(0) = 35 : arrPercent(1) = 25 : arrPercent(2) = 20
    arrPercent(3) = 12 : arrPercent(4) =  8

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
    objSheet.Cells(1, 1).Value = "資源項目"
    objSheet.Cells(1, 2).Value = "佔比（%）"

    With objSheet.Range("A1:B1")
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 4
        objSheet.Cells(i + 2, 1).Value = arrItems(i)
        objSheet.Cells(i + 2, 2).Value = arrPercent(i)
    Next i

    objSheet.Columns("A:B").AutoFit

    ' ── 插入 3D 圓餅圖 ───────────────────────────────────────────
    Set objChartObj = objSheet.ChartObjects.Add(200, 20, 420, 310)
    Set objChart    = objChartObj.Chart

    objChart.ChartType = xl3DPie
    objChart.SetSourceData objSheet.Range("A1:B6")

    ' ── 圖表格式設定 ────────────────────────────────────────────
    objChart.HasTitle        = True
    objChart.ChartTitle.Text = CHART_TITLE
    objChart.ChartTitle.Font.Size = 14
    objChart.ChartTitle.Font.Bold = True

    ' 顯示百分比 + 類別名稱
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
