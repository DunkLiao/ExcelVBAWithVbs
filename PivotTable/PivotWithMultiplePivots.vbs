' ============================================================
' PivotWithMultiplePivots.vbs
' 說明：使用 VBScript 在同一活頁簿建立兩個共用快取的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「人才招募」工作表填入招募示範資料
'   3. 建立共用同一 PivotCache 的兩個樞紐分析表：
'      - 樞紐1（列=部門，值=招募人數加總）
'      - 樞紐2（列=招募管道，欄=部門，值=招募費用加總）
'   4. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithMultiplePivots.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA   = "人才招募"
Const SHEET_PIVOT  = "樞紐分析表"
Const PIVOT_NAME1  = "部門招募人數"
Const PIVOT_NAME2  = "管道費用分析"
Const OUTPUT_FILE  = "18_PivotWithMultiplePivots.xlsx"

Const xlDatabase    = 1
Const xlRowField    = 1
Const xlColumnField = 2
Const xlDataField   = 3
Const xlSum         = -4157

' ── 範例資料（部門、招募管道、招募人數、招募費用）──────────
Dim arrDepts(19)
Dim arrChannels(19)
Dim arrCounts(19)
Dim arrCosts(19)

arrDepts(0)  = "研發部" : arrChannels(0)  = "人力銀行"  : arrCounts(0)  = 3 : arrCosts(0)  = 45000
arrDepts(1)  = "研發部" : arrChannels(1)  = "校園徵才"  : arrCounts(1)  = 5 : arrCosts(1)  = 30000
arrDepts(2)  = "研發部" : arrChannels(2)  = "獵頭公司"  : arrCounts(2)  = 2 : arrCosts(2)  = 120000
arrDepts(3)  = "研發部" : arrChannels(3)  = "員工推薦"  : arrCounts(3)  = 4 : arrCosts(3)  = 20000
arrDepts(4)  = "業務部" : arrChannels(4)  = "人力銀行"  : arrCounts(4)  = 6 : arrCosts(4)  = 54000
arrDepts(5)  = "業務部" : arrChannels(5)  = "校園徵才"  : arrCounts(5)  = 3 : arrCosts(5)  = 18000
arrDepts(6)  = "業務部" : arrChannels(6)  = "獵頭公司"  : arrCounts(6)  = 1 : arrCosts(6)  = 80000
arrDepts(7)  = "業務部" : arrChannels(7)  = "員工推薦"  : arrCounts(7)  = 8 : arrCosts(7)  = 40000
arrDepts(8)  = "生產部" : arrChannels(8)  = "人力銀行"  : arrCounts(8)  = 12: arrCosts(8)  = 72000
arrDepts(9)  = "生產部" : arrChannels(9)  = "校園徵才"  : arrCounts(9)  = 8 : arrCosts(9)  = 48000
arrDepts(10) = "生產部" : arrChannels(10) = "勞務派遣"  : arrCounts(10) = 20: arrCosts(10) = 60000
arrDepts(11) = "生產部" : arrChannels(11) = "員工推薦"  : arrCounts(11) = 5 : arrCosts(11) = 25000
arrDepts(12) = "行政部" : arrChannels(12) = "人力銀行"  : arrCounts(12) = 2 : arrCosts(12) = 18000
arrDepts(13) = "行政部" : arrChannels(13) = "員工推薦"  : arrCounts(13) = 3 : arrCosts(13) = 15000
arrDepts(14) = "財務部" : arrChannels(14) = "人力銀行"  : arrCounts(14) = 2 : arrCosts(14) = 24000
arrDepts(15) = "財務部" : arrChannels(15) = "獵頭公司"  : arrCounts(15) = 1 : arrCosts(15) = 70000
arrDepts(16) = "資訊部" : arrChannels(16) = "人力銀行"  : arrCounts(16) = 4 : arrCosts(16) = 48000
arrDepts(17) = "資訊部" : arrChannels(17) = "獵頭公司"  : arrCounts(17) = 2 : arrCosts(17) = 140000
arrDepts(18) = "資訊部" : arrChannels(18) = "校園徵才"  : arrCounts(18) = 6 : arrCosts(18) = 36000
arrDepts(19) = "資訊部" : arrChannels(19) = "員工推薦"  : arrCounts(19) = 3 : arrCosts(19) = 15000

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objDataSheet, objPivotSheet
Dim objCache, objPivot1, objPivot2, objField
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
objDataSheet.Cells(1, 1).Value = "部門"
objDataSheet.Cells(1, 2).Value = "招募管道"
objDataSheet.Cells(1, 3).Value = "招募人數"
objDataSheet.Cells(1, 4).Value = "招募費用"

With objDataSheet.Range("A1:D1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 19
    objDataSheet.Cells(i + 2, 1).Value = arrDepts(i)
    objDataSheet.Cells(i + 2, 2).Value = arrChannels(i)
    objDataSheet.Cells(i + 2, 3).Value = arrCounts(i)
    objDataSheet.Cells(i + 2, 4).Value = arrCosts(i)
Next

objDataSheet.Columns("A:D").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立共用的 PivotCache ─────────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:D21"))

' ────────────────────────────────────────────────────────────
' 樞紐1：部門招募人數（放置於 A3）
' ────────────────────────────────────────────────────────────
Set objPivot1 = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME1)

Set objField = objPivot1.PivotFields("部門")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot1.PivotFields("招募人數")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 招募人數"

Set objField = objPivot1.PivotFields("招募費用")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 招募費用"

' ────────────────────────────────────────────────────────────
' 樞紐2：招募管道 × 部門費用分析（放置於 H3）
' ────────────────────────────────────────────────────────────
Set objPivot2 = objCache.CreatePivotTable(objPivotSheet.Range("H3"), PIVOT_NAME2)

Set objField = objPivot2.PivotFields("招募管道")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot2.PivotFields("部門")
objField.Orientation = xlColumnField
objField.Position    = 1

Set objField = objPivot2.PivotFields("招募費用")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "費用合計"

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "多樞紐分析表：同一快取建立兩個樞紐（招募人數 & 管道費用分析）"
With objPivotSheet.Range("A1")
    .Font.Bold = True
    .Font.Size = 13
End With

' ── 儲存並關閉 ──────────────────────────────────────────────
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

Set objField      = Nothing
Set objPivot2     = Nothing
Set objPivot1     = Nothing
Set objCache      = Nothing
Set objPivotSheet = Nothing
Set objDataSheet  = Nothing
Set objWorkbook   = Nothing
Set objExcel      = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
