' ============================================================
' PivotWithChart.vbs
' 說明：使用 VBScript 自動建立樞紐分析表並搭配樞紐分析圖
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「季度銷售」工作表填入季度銷售示範資料
'   3. 建立樞紐分析表（列=地區，欄=季度，值=銷售額加總）
'   4. 在同一工作表插入嵌入式樞紐分析圖（群組直條圖）
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "季度銷售"
Const SHEET_PIVOT = "樞紐分析圖"
Const PIVOT_NAME  = "季度銷售樞紐"
Const OUTPUT_FILE = "10_PivotWithChart.xlsx"

Const xlDatabase         = 1
Const xlRowField         = 1
Const xlColumnField      = 2
Const xlDataField        = 3
Const xlSum              = -4157
Const xlClusteredColumn  = 51   ' 群組直條圖

' ── 範例資料（地區、季度、銷售額）──────────────────────────
Dim arrRegions(15)
Dim arrQtrs(15)
Dim arrAmounts(15)

arrRegions(0)  = "北區" : arrQtrs(0)  = "Q1" : arrAmounts(0)  = 120000
arrRegions(1)  = "北區" : arrQtrs(1)  = "Q2" : arrAmounts(1)  = 145000
arrRegions(2)  = "北區" : arrQtrs(2)  = "Q3" : arrAmounts(2)  = 168000
arrRegions(3)  = "北區" : arrQtrs(3)  = "Q4" : arrAmounts(3)  = 195000
arrRegions(4)  = "南區" : arrQtrs(4)  = "Q1" : arrAmounts(4)  = 98000
arrRegions(5)  = "南區" : arrQtrs(5)  = "Q2" : arrAmounts(5)  = 112000
arrRegions(6)  = "南區" : arrQtrs(6)  = "Q3" : arrAmounts(6)  = 134000
arrRegions(7)  = "南區" : arrQtrs(7)  = "Q4" : arrAmounts(7)  = 158000
arrRegions(8)  = "東區" : arrQtrs(8)  = "Q1" : arrAmounts(8)  = 87000
arrRegions(9)  = "東區" : arrQtrs(9)  = "Q2" : arrAmounts(9)  = 103000
arrRegions(10) = "東區" : arrQtrs(10) = "Q3" : arrAmounts(10) = 121000
arrRegions(11) = "東區" : arrQtrs(11) = "Q4" : arrAmounts(11) = 145000
arrRegions(12) = "西區" : arrQtrs(12) = "Q1" : arrAmounts(12) = 76000
arrRegions(13) = "西區" : arrQtrs(13) = "Q2" : arrAmounts(13) = 89000
arrRegions(14) = "西區" : arrQtrs(14) = "Q3" : arrAmounts(14) = 105000
arrRegions(15) = "西區" : arrQtrs(15) = "Q4" : arrAmounts(15) = 127000

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objDataSheet, objPivotSheet
Dim objCache, objPivot, objField
Dim objChartObj, objChart
Dim savePath, objShell, i

Set objShell = CreateObject("WScript.Shell")
savePath = objShell.SpecialFolders("Desktop") & "\" & OUTPUT_FILE
Set objShell = Nothing

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible       = False
objExcel.DisplayAlerts = False

Set objWorkbook   = objExcel.Workbooks.Add()
Set objDataSheet  = objWorkbook.Sheets(1)
objDataSheet.Name = SHEET_DATA

' ── 寫入標題列 ──────────────────────────────────────────────
objDataSheet.Cells(1, 1).Value = "地區"
objDataSheet.Cells(1, 2).Value = "季度"
objDataSheet.Cells(1, 3).Value = "銷售額"

With objDataSheet.Range("A1:C1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 15
    objDataSheet.Cells(i + 2, 1).Value = arrRegions(i)
    objDataSheet.Cells(i + 2, 2).Value = arrQtrs(i)
    objDataSheet.Cells(i + 2, 3).Value = arrAmounts(i)
Next

objDataSheet.Columns("A:C").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C17"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列、欄、值欄位 ──────────────────────────────────────
Set objField = objPivot.PivotFields("地區")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("季度")
objField.Orientation = xlColumnField
objField.Position    = 1

Set objField = objPivot.PivotFields("銷售額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 銷售額"

' ── 插入嵌入式樞紐分析圖 ─────────────────────────────────────
' 在樞紐分析表右方插入群組直條圖，資料來源設為樞紐分析表範圍
Set objChartObj = objPivotSheet.ChartObjects.Add(340, 50, 500, 320)
Set objChart    = objChartObj.Chart

objChart.SetSourceData objPivot.TableRange1
objChart.ChartType = xlClusteredColumn

' ── 設定圖表標題 ─────────────────────────────────────────────
objChart.HasTitle        = True
objChart.ChartTitle.Text = "2025 年各地區季度銷售額"
With objChart.ChartTitle.Font
    .Size = 13
    .Bold = True
End With

objChart.HasLegend = True

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "樞紐分析圖：依樞紐分析表資料自動生成群組直條圖"
With objPivotSheet.Range("A1")
    .Font.Bold = True
    .Font.Size = 14
End With

' ── 儲存並關閉 ──────────────────────────────────────────────
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

Set objChart      = Nothing
Set objChartObj   = Nothing
Set objField      = Nothing
Set objPivot      = Nothing
Set objCache      = Nothing
Set objPivotSheet = Nothing
Set objDataSheet  = Nothing
Set objWorkbook   = Nothing
Set objExcel      = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
