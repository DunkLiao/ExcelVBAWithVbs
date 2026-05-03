' ============================================================
' PivotWithRankNumber.vbs
' 說明：使用 VBScript 自動建立以排名數字顯示的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「業績排名」工作表填入業務員業績示範資料
'   3. 建立樞紐分析表（列=業務員，值=業績加總）
'   4. 新增第二個值欄位，改以「排名（由大到小）」方式顯示
'   5. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithRankNumber.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "業績排名"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "排名樞紐"
Const OUTPUT_FILE = "19_PivotWithRankNumber.xlsx"

Const xlDatabase       = 1
Const xlRowField       = 1
Const xlColumnField    = 2
Const xlDataField      = 3
Const xlSum            = -4157
Const xlRankDecreasing = 15   ' 由大到小排名（值越大排名越前）

' ── 範例資料（部門、業務員、業績金額）──────────────────────
Dim arrDepts(19)
Dim arrEmps(19)
Dim arrAmounts(19)

arrDepts(0)  = "A組" : arrEmps(0)  = "王志明" : arrAmounts(0)  = 342000
arrDepts(1)  = "A組" : arrEmps(1)  = "李雅琴" : arrAmounts(1)  = 218000
arrDepts(2)  = "A組" : arrEmps(2)  = "張建宏" : arrAmounts(2)  = 175000
arrDepts(3)  = "A組" : arrEmps(3)  = "陳淑芬" : arrAmounts(3)  = 290000
arrDepts(4)  = "A組" : arrEmps(4)  = "林冠廷" : arrAmounts(4)  = 156000
arrDepts(5)  = "B組" : arrEmps(5)  = "黃怡婷" : arrAmounts(5)  = 412000
arrDepts(6)  = "B組" : arrEmps(6)  = "吳俊男" : arrAmounts(6)  = 267000
arrDepts(7)  = "B組" : arrEmps(7)  = "鄭麗華" : arrAmounts(7)  = 198000
arrDepts(8)  = "B組" : arrEmps(8)  = "謝文成" : arrAmounts(8)  = 325000
arrDepts(9)  = "B組" : arrEmps(9)  = "許志豪" : arrAmounts(9)  = 143000
arrDepts(10) = "C組" : arrEmps(10) = "楊雅雯" : arrAmounts(10) = 378000
arrDepts(11) = "C組" : arrEmps(11) = "蔡明宏" : arrAmounts(11) = 231000
arrDepts(12) = "C組" : arrEmps(12) = "洪淑慧" : arrAmounts(12) = 185000
arrDepts(13) = "C組" : arrEmps(13) = "劉建國" : arrAmounts(13) = 298000
arrDepts(14) = "C組" : arrEmps(14) = "賴怡君" : arrAmounts(14) = 124000
arrDepts(15) = "D組" : arrEmps(15) = "葉士豪" : arrAmounts(15) = 456000
arrDepts(16) = "D組" : arrEmps(16) = "方思穎" : arrAmounts(16) = 312000
arrDepts(17) = "D組" : arrEmps(17) = "周文傑" : arrAmounts(17) = 189000
arrDepts(18) = "D組" : arrEmps(18) = "江美玲" : arrAmounts(18) = 267000
arrDepts(19) = "D組" : arrEmps(19) = "莊裕民" : arrAmounts(19) = 98000

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objDataSheet, objPivotSheet
Dim objCache, objPivot, objField, objDataField
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
objDataSheet.Cells(1, 1).Value = "組別"
objDataSheet.Cells(1, 2).Value = "業務員"
objDataSheet.Cells(1, 3).Value = "業績金額"

With objDataSheet.Range("A1:C1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 19
    objDataSheet.Cells(i + 2, 1).Value = arrDepts(i)
    objDataSheet.Cells(i + 2, 2).Value = arrEmps(i)
    objDataSheet.Cells(i + 2, 3).Value = arrAmounts(i)
Next

objDataSheet.Columns("A:C").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C21"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列、欄、值欄位 ──────────────────────────────────────
Set objField = objPivot.PivotFields("業務員")
objField.Orientation = xlRowField
objField.Position    = 1

' 第一個值欄位：業績金額加總
Set objField = objPivot.PivotFields("業績金額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "業績金額"

' 第二個值欄位：排名（由大到小）
Set objField = objPivot.PivotFields("業績金額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "全體排名"

' ── 設定排名計算方式（依業務員欄位，由大到小排名）───────────
Set objDataField = objPivot.DataFields("全體排名")
objDataField.Calculation = xlRankDecreasing
objDataField.BaseField   = "業務員"

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "排名顯示樞紐分析表：業績金額旁同時顯示各業務員的全體排名"
With objPivotSheet.Range("A1")
    .Font.Bold = True
    .Font.Size = 14
End With

' ── 儲存並關閉 ──────────────────────────────────────────────
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

Set objDataField  = Nothing
Set objField      = Nothing
Set objPivot      = Nothing
Set objCache      = Nothing
Set objPivotSheet = Nothing
Set objDataSheet  = Nothing
Set objWorkbook   = Nothing
Set objExcel      = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
