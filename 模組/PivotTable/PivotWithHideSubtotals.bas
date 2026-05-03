Attribute VB_Name = "PivotWithHideSubtotals"
' ============================================================
' PivotWithHideSubtotals.bas
' 說明：使用 Excel VBA 自動建立隱藏所有小計的樞紐分析表
' 功能：
'   1. 建立新活頁簿
'   2. 在「薪資資料」工作表填入人事薪資示範資料
'   3. 建立含兩層列欄位的樞紐分析表
'   4. 隱藏所有列欄位的小計列，使報表更簡潔
'   5. 同時隱藏總計列
'   6. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 PivotWithHideSubtotals 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  As String = "薪資資料"
Const SHEET_PIVOT As String = "樞紐分析表"
Const PIVOT_NAME  As String = "隱藏小計樞紐"
Const OUTPUT_FILE As String = "16_PivotWithHideSubtotals.xlsx"

Sub PivotWithHideSubtotals()

    ' ── 範例資料（公司別、部門、員工人數、薪資總額）────────────
    Dim arrCompanies(19)  As String
    Dim arrDepts(19)      As String
    Dim arrHeadcounts(19) As Long
    Dim arrSalaries(19)   As Long

    arrCompanies(0)  = "台灣總公司" : arrDepts(0)  = "研發部" : arrHeadcounts(0)  = 25 : arrSalaries(0)  = 2250000
    arrCompanies(1)  = "台灣總公司" : arrDepts(1)  = "業務部" : arrHeadcounts(1)  = 18 : arrSalaries(1)  = 1440000
    arrCompanies(2)  = "台灣總公司" : arrDepts(2)  = "生產部" : arrHeadcounts(2)  = 42 : arrSalaries(2)  = 2940000
    arrCompanies(3)  = "台灣總公司" : arrDepts(3)  = "行政部" : arrHeadcounts(3)  = 10 : arrSalaries(3)  = 680000
    arrCompanies(4)  = "台灣總公司" : arrDepts(4)  = "財務部" : arrHeadcounts(4)  = 8  : arrSalaries(4)  = 640000
    arrCompanies(5)  = "台中分公司" : arrDepts(5)  = "研發部" : arrHeadcounts(5)  = 12 : arrSalaries(5)  = 1020000
    arrCompanies(6)  = "台中分公司" : arrDepts(6)  = "業務部" : arrHeadcounts(6)  = 15 : arrSalaries(6)  = 1125000
    arrCompanies(7)  = "台中分公司" : arrDepts(7)  = "生產部" : arrHeadcounts(7)  = 35 : arrSalaries(7)  = 2275000
    arrCompanies(8)  = "台中分公司" : arrDepts(8)  = "行政部" : arrHeadcounts(8)  = 6  : arrSalaries(8)  = 390000
    arrCompanies(9)  = "台中分公司" : arrDepts(9)  = "財務部" : arrHeadcounts(9)  = 5  : arrSalaries(9)  = 375000
    arrCompanies(10) = "高雄分公司" : arrDepts(10) = "研發部" : arrHeadcounts(10) = 8  : arrSalaries(10) = 680000
    arrCompanies(11) = "高雄分公司" : arrDepts(11) = "業務部" : arrHeadcounts(11) = 20 : arrSalaries(11) = 1400000
    arrCompanies(12) = "高雄分公司" : arrDepts(12) = "生產部" : arrHeadcounts(12) = 55 : arrSalaries(12) = 3300000
    arrCompanies(13) = "高雄分公司" : arrDepts(13) = "行政部" : arrHeadcounts(13) = 7  : arrSalaries(13) = 455000
    arrCompanies(14) = "高雄分公司" : arrDepts(14) = "財務部" : arrHeadcounts(14) = 4  : arrSalaries(14) = 300000
    arrCompanies(15) = "新加坡子公司" : arrDepts(15) = "研發部" : arrHeadcounts(15) = 15 : arrSalaries(15) = 1875000
    arrCompanies(16) = "新加坡子公司" : arrDepts(16) = "業務部" : arrHeadcounts(16) = 10 : arrSalaries(16) = 1150000
    arrCompanies(17) = "新加坡子公司" : arrDepts(17) = "生產部" : arrHeadcounts(17) = 0  : arrSalaries(17) = 0
    arrCompanies(18) = "新加坡子公司" : arrDepts(18) = "行政部" : arrHeadcounts(18) = 5  : arrSalaries(18) = 525000
    arrCompanies(19) = "新加坡子公司" : arrDepts(19) = "財務部" : arrHeadcounts(19) = 3  : arrSalaries(19) = 345000

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook      As Workbook
    Dim objDataSheet     As Worksheet
    Dim objPivotSheet    As Worksheet
    Dim objCache         As PivotCache
    Dim objPivot         As PivotTable
    Dim objField         As PivotField
    Dim arrNoSubtotals(11) As Boolean
    Dim savePath         As String
    Dim i                As Integer
    Dim j                As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook   = Workbooks.Add()
    Set objDataSheet  = objWorkbook.Sheets(1)
    objDataSheet.Name = SHEET_DATA

    ' ── 寫入標題列 ──────────────────────────────────────────────
    objDataSheet.Cells(1, 1).Value = "公司別"
    objDataSheet.Cells(1, 2).Value = "部門"
    objDataSheet.Cells(1, 3).Value = "員工人數"
    objDataSheet.Cells(1, 4).Value = "薪資總額"

    With objDataSheet.Range("A1:D1")
        .Font.Bold           = True
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 19
        objDataSheet.Cells(i + 2, 1).Value = arrCompanies(i)
        objDataSheet.Cells(i + 2, 2).Value = arrDepts(i)
        objDataSheet.Cells(i + 2, 3).Value = arrHeadcounts(i)
        objDataSheet.Cells(i + 2, 4).Value = arrSalaries(i)
    Next i

    objDataSheet.Columns("A:D").AutoFit

    ' ── 新增樞紐分析表工作表 ─────────────────────────────────────
    Set objPivotSheet  = objWorkbook.Sheets.Add()
    objPivotSheet.Name = SHEET_PIVOT
    objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

    ' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
    Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:D21"))
    Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

    ' ── 設定列欄位 ──────────────────────────────────────────────
    Set objField = objPivot.PivotFields("公司別")
    objField.Orientation = xlRowField
    objField.Position    = 1

    Set objField = objPivot.PivotFields("部門")
    objField.Orientation = xlRowField
    objField.Position    = 2

    ' ── 設定值欄位（員工人數、薪資總額）────────────────────────
    Set objField = objPivot.PivotFields("員工人數")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "加總 - 員工人數"

    Set objField = objPivot.PivotFields("薪資總額")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "加總 - 薪資總額"

    ' ── 隱藏所有列欄位的小計 ────────────────────────────────────
    ' Subtotals 屬性是長度為 12 的布林陣列，全設為 False 即隱藏小計
    For j = 0 To 11
        arrNoSubtotals(j) = False
    Next j

    objPivot.PivotFields("公司別").Subtotals = arrNoSubtotals
    objPivot.PivotFields("部門").Subtotals   = arrNoSubtotals

    ' ── 隱藏欄總計（列方向的總計列）────────────────────────────
    objPivot.ColumnGrand = False

    ' ── 加入說明標題 ─────────────────────────────────────────────
    objPivotSheet.Range("A1").Value = "隱藏小計樞紐分析表：移除所有中間小計列，報表更簡潔"
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
