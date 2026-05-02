' ============================================================
' CreateRadarChart.vbs
' 說明：使用 VBScript 自動建立 Excel 雷達圖範例
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在工作表填入示範資料（員工技能評分）
'   3. 插入雷達圖
'   4. 設定圖表標題、圖例等格式
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript charts\CreateRadarChart.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const CHART_TITLE = "員工技能雷達圖"
Const SHEET_NAME  = "技能評分"
Const OUTPUT_FILE = "RadarChartExample.xlsx"

' xlRadar = -4151（雷達圖）
Const xlRadar = -4151

' ── 範例資料 ────────────────────────────────────────────────
' 欄位：技能維度, 員工A, 員工B
Dim arrSkills(5)
arrSkills(0) = "溝通能力" : arrSkills(1) = "技術能力" : arrSkills(2) = "創新思維"
arrSkills(3) = "團隊合作" : arrSkills(4) = "問題解決" : arrSkills(5) = "時間管理"

Dim arrEmpA(5)
arrEmpA(0) = 85 : arrEmpA(1) = 92 : arrEmpA(2) = 78
arrEmpA(3) = 88 : arrEmpA(4) = 90 : arrEmpA(5) = 75

Dim arrEmpB(5)
arrEmpB(0) = 90 : arrEmpB(1) = 75 : arrEmpB(2) = 88
arrEmpB(3) = 82 : arrEmpB(4) = 85 : arrEmpB(5) = 92

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objSheet, objChartObj, objChart
Dim savePath, objShell, i

Set objShell = CreateObject("WScript.Shell")
savePath = objShell.SpecialFolders("Desktop") & "\" & OUTPUT_FILE
Set objShell = Nothing

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible       = False
objExcel.DisplayAlerts = False

Set objWorkbook = objExcel.Workbooks.Add()
Set objSheet    = objWorkbook.Sheets(1)
objSheet.Name   = SHEET_NAME

' ── 寫入標題列 ──────────────────────────────────────────────
objSheet.Cells(1, 1).Value = "技能項目"
objSheet.Cells(1, 2).Value = "員工A"
objSheet.Cells(1, 3).Value = "員工B"

With objSheet.Range("A1:C1")
    .Font.Bold           = True
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 5
    objSheet.Cells(i + 2, 1).Value = arrSkills(i)
    objSheet.Cells(i + 2, 2).Value = arrEmpA(i)
    objSheet.Cells(i + 2, 3).Value = arrEmpB(i)
Next

objSheet.Columns("A:C").AutoFit()

' ── 插入雷達圖 ───────────────────────────────────────────────
Set objChartObj = objSheet.ChartObjects.Add(230, 20, 400, 320)
Set objChart    = objChartObj.Chart

objChart.ChartType = xlRadar
objChart.SetSourceData objSheet.Range("A1:C7")

' ── 圖表格式設定 ────────────────────────────────────────────
objChart.HasTitle        = True
objChart.ChartTitle.Text = CHART_TITLE
objChart.ChartTitle.Font.Size = 14
objChart.ChartTitle.Font.Bold = True

objChart.HasLegend = True

' ── 儲存並關閉 ──────────────────────────────────────────────
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

Set objChart    = Nothing
Set objChartObj = Nothing
Set objSheet    = Nothing
Set objWorkbook = Nothing
Set objExcel    = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
