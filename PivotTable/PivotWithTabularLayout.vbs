' ============================================================
' PivotWithTabularLayout.vbs
' 說明：使用 VBScript 自動建立表格式版面配置的樞紐分析表
' 功能：
'   1. 開啟 Excel 並建立新活頁簿
'   2. 在「採購明細」工作表填入採購示範資料
'   3. 建立樞紐分析表（含兩層列欄位）
'   4. 將所有列欄位設定為「表格式版面配置」（每欄獨立一欄，方便閱讀）
'   5. 關閉所有列欄位的小計（表格式慣用方式）
'   6. 將結果儲存至桌面
' 執行方式：在命令提示字元輸入  cscript PivotTable\PivotWithTabularLayout.vbs
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  = "採購明細"
Const SHEET_PIVOT = "樞紐分析表"
Const PIVOT_NAME  = "表格式版面樞紐"
Const OUTPUT_FILE = "15_PivotWithTabularLayout.xlsx"

Const xlDatabase    = 1
Const xlRowField    = 1
Const xlColumnField = 2
Const xlDataField   = 3
Const xlSum         = -4157
Const xlTabular     = 1    ' 表格式版面（xlLayoutFormType）
Const xlAtBottom    = 2    ' 小計位置（顯示於底部）

' ── 範例資料（供應商類別、供應商、採購品項、採購金額）──────
Dim arrCats(19)
Dim arrVendors(19)
Dim arrItems(19)
Dim arrAmounts(19)

arrCats(0)  = "原物料" : arrVendors(0)  = "鋼鐵大廠A" : arrItems(0)  = "不鏽鋼板" : arrAmounts(0)  = 320000
arrCats(1)  = "原物料" : arrVendors(1)  = "鋼鐵大廠A" : arrItems(1)  = "鋁合金棒" : arrAmounts(1)  = 185000
arrCats(2)  = "原物料" : arrVendors(2)  = "化工原料B" : arrItems(2)  = "環氧樹脂" : arrAmounts(2)  = 92000
arrCats(3)  = "原物料" : arrVendors(3)  = "化工原料B" : arrItems(3)  = "固化劑"   : arrAmounts(3)  = 67000
arrCats(4)  = "原物料" : arrVendors(4)  = "化工原料B" : arrItems(4)  = "溶劑"     : arrAmounts(4)  = 45000
arrCats(5)  = "零組件" : arrVendors(5)  = "電子零件C" : arrItems(5)  = "電阻"     : arrAmounts(5)  = 28000
arrCats(6)  = "零組件" : arrVendors(6)  = "電子零件C" : arrItems(6)  = "電容"     : arrAmounts(6)  = 35000
arrCats(7)  = "零組件" : arrVendors(7)  = "電子零件C" : arrItems(7)  = "IC 晶片"  : arrAmounts(7)  = 248000
arrCats(8)  = "零組件" : arrVendors(8)  = "機械零件D" : arrItems(8)  = "軸承"     : arrAmounts(8)  = 76000
arrCats(9)  = "零組件" : arrVendors(9)  = "機械零件D" : arrItems(9)  = "齒輪"     : arrAmounts(9)  = 112000
arrCats(10) = "零組件" : arrVendors(10) = "機械零件D" : arrItems(10) = "油封"     : arrAmounts(10) = 38000
arrCats(11) = "耗材"   : arrVendors(11) = "辦公耗材E" : arrItems(11) = "紙張"     : arrAmounts(11) = 15000
arrCats(12) = "耗材"   : arrVendors(12) = "辦公耗材E" : arrItems(12) = "墨水匣"   : arrAmounts(12) = 22000
arrCats(13) = "耗材"   : arrVendors(13) = "辦公耗材E" : arrItems(13) = "清潔用品" : arrAmounts(13) = 18000
arrCats(14) = "耗材"   : arrVendors(14) = "工業耗材F" : arrItems(14) = "砂紙"     : arrAmounts(14) = 9000
arrCats(15) = "耗材"   : arrVendors(15) = "工業耗材F" : arrItems(15) = "切削油"   : arrAmounts(15) = 31000
arrCats(16) = "設備"   : arrVendors(16) = "設備廠商G" : arrItems(16) = "工具機"   : arrAmounts(16) = 580000
arrCats(17) = "設備"   : arrVendors(17) = "設備廠商G" : arrItems(17) = "量測儀器" : arrAmounts(17) = 195000
arrCats(18) = "設備"   : arrVendors(18) = "設備廠商H" : arrItems(18) = "輸送帶"   : arrAmounts(18) = 142000
arrCats(19) = "設備"   : arrVendors(19) = "設備廠商H" : arrItems(19) = "防護設備" : arrAmounts(19) = 87000

' ── 主程式 ──────────────────────────────────────────────────
Dim objExcel, objWorkbook, objDataSheet, objPivotSheet
Dim objCache, objPivot, objField
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
objDataSheet.Cells(1, 1).Value = "採購類別"
objDataSheet.Cells(1, 2).Value = "供應商"
objDataSheet.Cells(1, 3).Value = "採購品項"
objDataSheet.Cells(1, 4).Value = "採購金額"

With objDataSheet.Range("A1:D1")
    .Font.Bold           = True
    .Interior.Color      = RGB(68, 114, 196)
    .Font.Color          = RGB(255, 255, 255)
    .HorizontalAlignment = -4108  ' xlCenter
End With

' ── 寫入資料列 ──────────────────────────────────────────────
For i = 0 To 19
    objDataSheet.Cells(i + 2, 1).Value = arrCats(i)
    objDataSheet.Cells(i + 2, 2).Value = arrVendors(i)
    objDataSheet.Cells(i + 2, 3).Value = arrItems(i)
    objDataSheet.Cells(i + 2, 4).Value = arrAmounts(i)
Next

objDataSheet.Columns("A:D").AutoFit()

' ── 新增樞紐分析表工作表 ─────────────────────────────────────
Set objPivotSheet  = objWorkbook.Sheets.Add()
objPivotSheet.Name = SHEET_PIVOT
objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:D21"))
Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

' ── 設定列欄位（採購類別 > 供應商 > 採購品項）───────────────
Set objField = objPivot.PivotFields("採購類別")
objField.Orientation = xlRowField
objField.Position    = 1

Set objField = objPivot.PivotFields("供應商")
objField.Orientation = xlRowField
objField.Position    = 2

Set objField = objPivot.PivotFields("採購品項")
objField.Orientation = xlRowField
objField.Position    = 3

' ── 設定值欄位 ──────────────────────────────────────────────
Set objField = objPivot.PivotFields("採購金額")
objField.Orientation = xlDataField
objField.Function    = xlSum
objField.Name        = "加總 - 採購金額"

' ── 將所有列欄位設定為表格式版面配置 ────────────────────────
' LayoutForm = xlTabular(1)：每個列欄位各自獨立一欄，橫向展開
' RepeatLabels = True：重複顯示群組標籤以便閱讀
Dim fldName
For Each fldName In Array("採購類別", "供應商", "採購品項")
    With objPivot.PivotFields(fldName)
        .LayoutForm              = xlTabular
        .RepeatLabels            = True
        .LayoutSubtotalLocation  = xlAtBottom
    End With
Next

' ── 加入說明標題 ─────────────────────────────────────────────
objPivotSheet.Range("A1").Value = "表格式版面配置樞紐分析表：各欄位獨立一欄，結構清晰易讀"
With objPivotSheet.Range("A1")
    .Font.Bold = True
    .Font.Size = 14
End With

' ── 儲存並關閉 ──────────────────────────────────────────────
objWorkbook.SaveAs savePath, 51
objWorkbook.Close False
objExcel.Quit

Set objField      = Nothing
Set objPivot      = Nothing
Set objCache      = Nothing
Set objPivotSheet = Nothing
Set objDataSheet  = Nothing
Set objWorkbook   = Nothing
Set objExcel      = Nothing

WScript.Echo "完成！檔案已儲存至：" & savePath
