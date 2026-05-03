Attribute VB_Name = "CreatePieExplodedChart"
' ============================================================
' CreatePieExplodedChart.bas
' 說明：使用 Excel VBA 自動建立分裂圓餅圖範例
' 功能：
'   1. 建立新活頁簿
'   2. 在工作表填入示範資料（生活費用類別佔比）
'   3. 插入分裂圓餅圖
'   4. 設定圖表標題、百分比資料標籤等格式
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 CreatePieExplodedChart 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE As String = "月度生活費用類別佔比"
Const SHEET_NAME  As String = "費用分析"
Const OUTPUT_FILE As String = "PieExplodedChartExample.xlsx"

Sub CreatePieExplodedChart()

    ' ── 範例資料 ────────────────────────────────────────────────
    Dim arrCategories(5) As String
    arrCategories(0) = "飲食" : arrCategories(1) = "租房" : arrCategories(2) = "交通"
    arrCategories(3) = "娛樂" : arrCategories(4) = "醫療" : arrCategories(5) = "其他"

    Dim arrAmount(5) As Long
    arrAmount(0) = 8000 : arrAmount(1) = 15000 : arrAmount(2) = 3000
    arrAmount(3) = 2500 : arrAmount(4) = 1500  : arrAmount(5) = 2000

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
    objSheet.Cells(1, 1).Value = "費用類別"
    objSheet.Cells(1, 2).Value = "金額（元）"

    With objSheet.Range("A1:B1")
        .Font.Bold           = True
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 5
        objSheet.Cells(i + 2, 1).Value = arrCategories(i)
        objSheet.Cells(i + 2, 2).Value = arrAmount(i)
    Next i

    objSheet.Columns("A:B").AutoFit

    ' ── 插入分裂圓餅圖 ───────────────────────────────────────────
    Set objChartObj = objSheet.ChartObjects.Add(200, 20, 420, 310)
    Set objChart    = objChartObj.Chart

    objChart.ChartType = xlPieExploded
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
