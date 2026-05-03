Attribute VB_Name = "PivotWithTabularLayout"
' ============================================================
' PivotWithTabularLayout.bas
' 說明：使用 Excel VBA 自動建立套用表格式版面配置的樞紐分析表
' 功能：
'   1. 建立新活頁簿
'   2. 在「採購記錄」工作表填入採購示範資料
'   3. 建立含三層列欄位的樞紐分析表
'   4. 將每個列欄位設為表格式版面（LayoutForm），並啟用重複標籤
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 PivotWithTabularLayout 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  As String = "採購記錄"
Const SHEET_PIVOT As String = "樞紐分析表"
Const PIVOT_NAME  As String = "表格式版面樞紐"
Const OUTPUT_FILE As String = "15_PivotWithTabularLayout.xlsx"

Sub PivotWithTabularLayout()

    ' ── 範例資料（採購類別、供應商、採購品項、採購金額）────────
    Dim arrCats(19)     As String
    Dim arrVendors(19)  As String
    Dim arrItems(19)    As String
    Dim arrAmounts(19)  As Long

    arrCats(0)  = "辦公用品" : arrVendors(0)  = "供應商A" : arrItems(0)  = "A4 影印紙"     : arrAmounts(0)  = 12000
    arrCats(1)  = "辦公用品" : arrVendors(1)  = "供應商A" : arrItems(1)  = "原子筆（箱）" : arrAmounts(1)  = 3500
    arrCats(2)  = "辦公用品" : arrVendors(2)  = "供應商A" : arrItems(2)  = "印表機碳粉"   : arrAmounts(2)  = 8800
    arrCats(3)  = "辦公用品" : arrVendors(3)  = "供應商A" : arrItems(3)  = "文件夾（盒）" : arrAmounts(3)  = 2100
    arrCats(4)  = "辦公用品" : arrVendors(4)  = "供應商A" : arrItems(4)  = "剪刀尺規組"   : arrAmounts(4)  = 1800
    arrCats(5)  = "電腦設備" : arrVendors(5)  = "供應商B" : arrItems(5)  = "桌上型電腦"   : arrAmounts(5)  = 350000
    arrCats(6)  = "電腦設備" : arrVendors(6)  = "供應商B" : arrItems(6)  = "27吋螢幕"      : arrAmounts(6)  = 75000
    arrCats(7)  = "電腦設備" : arrVendors(7)  = "供應商B" : arrItems(7)  = "機械鍵盤"      : arrAmounts(7)  = 18000
    arrCats(8)  = "電腦設備" : arrVendors(8)  = "供應商B" : arrItems(8)  = "無線滑鼠"      : arrAmounts(8)  = 9600
    arrCats(9)  = "電腦設備" : arrVendors(9)  = "供應商B" : arrItems(9)  = "網路交換器"    : arrAmounts(9)  = 42000
    arrCats(10) = "清潔用品" : arrVendors(10) = "供應商C" : arrItems(10) = "洗手乳（桶）" : arrAmounts(10) = 4200
    arrCats(11) = "清潔用品" : arrVendors(11) = "供應商C" : arrItems(11) = "酒精噴霧"      : arrAmounts(11) = 3600
    arrCats(12) = "清潔用品" : arrVendors(12) = "供應商C" : arrItems(12) = "衛生紙（箱）" : arrAmounts(12) = 2800
    arrCats(13) = "清潔用品" : arrVendors(13) = "供應商C" : arrItems(13) = "垃圾袋（捲）" : arrAmounts(13) = 1500
    arrCats(14) = "清潔用品" : arrVendors(14) = "供應商C" : arrItems(14) = "清潔劑（瓶）" : arrAmounts(14) = 980
    arrCats(15) = "飲料零食" : arrVendors(15) = "供應商D" : arrItems(15) = "礦泉水（箱）" : arrAmounts(15) = 5400
    arrCats(16) = "飲料零食" : arrVendors(16) = "供應商D" : arrItems(16) = "咖啡豆（包）" : arrAmounts(16) = 9800
    arrCats(17) = "飲料零食" : arrVendors(17) = "供應商D" : arrItems(17) = "茶包（盒）"   : arrAmounts(17) = 3200
    arrCats(18) = "飲料零食" : arrVendors(18) = "供應商D" : arrItems(18) = "餅乾（箱）"   : arrAmounts(18) = 4500
    arrCats(19) = "飲料零食" : arrVendors(19) = "供應商D" : arrItems(19) = "糖果（袋）"   : arrAmounts(19) = 1200

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook   As Workbook
    Dim objDataSheet  As Worksheet
    Dim objPivotSheet As Worksheet
    Dim objCache      As PivotCache
    Dim objPivot      As PivotTable
    Dim objField      As PivotField
    Dim fldName       As Variant
    Dim savePath      As String
    Dim i             As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook   = Workbooks.Add()
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
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 19
        objDataSheet.Cells(i + 2, 1).Value = arrCats(i)
        objDataSheet.Cells(i + 2, 2).Value = arrVendors(i)
        objDataSheet.Cells(i + 2, 3).Value = arrItems(i)
        objDataSheet.Cells(i + 2, 4).Value = arrAmounts(i)
    Next i

    objDataSheet.Columns("A:D").AutoFit

    ' ── 新增樞紐分析表工作表 ─────────────────────────────────────
    Set objPivotSheet  = objWorkbook.Sheets.Add()
    objPivotSheet.Name = SHEET_PIVOT
    objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

    ' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
    Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:D21"))
    Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

    ' ── 設定三層列欄位 ──────────────────────────────────────────
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

    ' ── 將每個列欄位設為表格式版面，並啟用重複標籤 ─────────────
    ' xlTabular = 2（表格式），LayoutSubtotalLocation 設為 xlAtBottom
    For Each fldName In Array("採購類別", "供應商", "採購品項")
        With objPivot.PivotFields(fldName)
            .LayoutForm             = xlTabular
            .RepeatLabels           = True
            .LayoutSubtotalLocation = xlAtBottom
        End With
    Next fldName

    ' ── 加入說明標題 ─────────────────────────────────────────────
    objPivotSheet.Range("A1").Value = "表格式版面樞紐分析表：採購類別 > 供應商 > 採購品項 三層列欄位"
    With objPivotSheet.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
    End With

    ' ── 儲存 ────────────────────────────────────────────────────
    objWorkbook.SaveAs savePath, xlOpenXMLWorkbook

    Set objField      = Nothing
    Set objPivot      = Nothing
    Set objCache      = Nothing
    Set objPivotSheet = Nothing
    Set objDataSheet  = Nothing
    Set objWorkbook   = Nothing

    MsgBox "完成！檔案已儲存至：" & savePath

End Sub
