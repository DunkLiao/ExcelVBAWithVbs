Attribute VB_Name = "PivotWithSlicer"
' ============================================================
' PivotWithSlicer.bas
' 說明：使用 Excel VBA 自動建立含交叉分析篩選器的樞紐分析表
' 功能：
'   1. 建立新活頁簿
'   2. 在「員工薪資」工作表填入薪資示範資料
'   3. 建立樞紐分析表（列=部門，欄=職級，值=薪資加總）
'   4. 新增「部門」與「職級」兩個交叉分析篩選器（Slicer）
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 PivotWithSlicer 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  As String = "員工薪資"
Const SHEET_PIVOT As String = "樞紐分析表"
Const PIVOT_NAME  As String = "篩選器樞紐"
Const OUTPUT_FILE As String = "09_PivotWithSlicer.xlsx"

Sub PivotWithSlicer()

    ' ── 範例資料（部門、職級、員工、薪資）──────────────────────
    Dim arrDepts(19)    As String
    Dim arrGrades(19)   As String
    Dim arrEmps(19)     As String
    Dim arrSalaries(19) As Long

    arrDepts(0)  = "研發部" : arrGrades(0)  = "高階" : arrEmps(0)  = "工程師A" : arrSalaries(0)  = 150000
    arrDepts(1)  = "研發部" : arrGrades(1)  = "高階" : arrEmps(1)  = "工程師B" : arrSalaries(1)  = 145000
    arrDepts(2)  = "研發部" : arrGrades(2)  = "中階" : arrEmps(2)  = "工程師C" : arrSalaries(2)  = 95000
    arrDepts(3)  = "研發部" : arrGrades(3)  = "中階" : arrEmps(3)  = "工程師D" : arrSalaries(3)  = 88000
    arrDepts(4)  = "研發部" : arrGrades(4)  = "初階" : arrEmps(4)  = "工程師E" : arrSalaries(4)  = 55000
    arrDepts(5)  = "業務部" : arrGrades(5)  = "高階" : arrEmps(5)  = "業務員A" : arrSalaries(5)  = 120000
    arrDepts(6)  = "業務部" : arrGrades(6)  = "高階" : arrEmps(6)  = "業務員B" : arrSalaries(6)  = 115000
    arrDepts(7)  = "業務部" : arrGrades(7)  = "中階" : arrEmps(7)  = "業務員C" : arrSalaries(7)  = 78000
    arrDepts(8)  = "業務部" : arrGrades(8)  = "中階" : arrEmps(8)  = "業務員D" : arrSalaries(8)  = 72000
    arrDepts(9)  = "業務部" : arrGrades(9)  = "初階" : arrEmps(9)  = "業務員E" : arrSalaries(9)  = 45000
    arrDepts(10) = "行政部" : arrGrades(10) = "高階" : arrEmps(10) = "行政員A" : arrSalaries(10) = 105000
    arrDepts(11) = "行政部" : arrGrades(11) = "中階" : arrEmps(11) = "行政員B" : arrSalaries(11) = 68000
    arrDepts(12) = "行政部" : arrGrades(12) = "中階" : arrEmps(12) = "行政員C" : arrSalaries(12) = 65000
    arrDepts(13) = "行政部" : arrGrades(13) = "初階" : arrEmps(13) = "行政員D" : arrSalaries(13) = 42000
    arrDepts(14) = "行政部" : arrGrades(14) = "初階" : arrEmps(14) = "行政員E" : arrSalaries(14) = 40000
    arrDepts(15) = "財務部" : arrGrades(15) = "高階" : arrEmps(15) = "財務員A" : arrSalaries(15) = 125000
    arrDepts(16) = "財務部" : arrGrades(16) = "高階" : arrEmps(16) = "財務員B" : arrSalaries(16) = 118000
    arrDepts(17) = "財務部" : arrGrades(17) = "中階" : arrEmps(17) = "財務員C" : arrSalaries(17) = 82000
    arrDepts(18) = "財務部" : arrGrades(18) = "中階" : arrEmps(18) = "財務員D" : arrSalaries(18) = 78000
    arrDepts(19) = "財務部" : arrGrades(19) = "初階" : arrEmps(19) = "財務員E" : arrSalaries(19) = 50000

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook      As Workbook
    Dim objDataSheet     As Worksheet
    Dim objPivotSheet    As Worksheet
    Dim objCache         As PivotCache
    Dim objPivot         As PivotTable
    Dim objField         As PivotField
    Dim objSlicerCache1  As SlicerCache
    Dim objSlicerCache2  As SlicerCache
    Dim savePath         As String
    Dim i                As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook   = Workbooks.Add()
    Set objDataSheet  = objWorkbook.Sheets(1)
    objDataSheet.Name = SHEET_DATA

    ' ── 寫入標題列 ──────────────────────────────────────────────
    objDataSheet.Cells(1, 1).Value = "部門"
    objDataSheet.Cells(1, 2).Value = "職級"
    objDataSheet.Cells(1, 3).Value = "員工"
    objDataSheet.Cells(1, 4).Value = "薪資"

    With objDataSheet.Range("A1:D1")
        .Font.Bold           = True
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 19
        objDataSheet.Cells(i + 2, 1).Value = arrDepts(i)
        objDataSheet.Cells(i + 2, 2).Value = arrGrades(i)
        objDataSheet.Cells(i + 2, 3).Value = arrEmps(i)
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

    ' ── 設定列、欄、值欄位 ──────────────────────────────────────
    Set objField = objPivot.PivotFields("部門")
    objField.Orientation = xlRowField
    objField.Position    = 1

    Set objField = objPivot.PivotFields("職級")
    objField.Orientation = xlColumnField
    objField.Position    = 1

    Set objField = objPivot.PivotFields("薪資")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "加總 - 薪資"

    ' ── 新增「部門」交叉分析篩選器 ──────────────────────────────
    Set objSlicerCache1 = objWorkbook.SlicerCaches.Add2(objPivot, "部門")
    objSlicerCache1.Slicers.Add objPivotSheet, , "部門篩選器", "部門", 20, 380, 144, 168

    ' ── 新增「職級」交叉分析篩選器 ──────────────────────────────
    Set objSlicerCache2 = objWorkbook.SlicerCaches.Add2(objPivot, "職級")
    objSlicerCache2.Slicers.Add objPivotSheet, , "職級篩選器", "職級", 240, 380, 144, 120

    ' ── 加入說明標題 ─────────────────────────────────────────────
    objPivotSheet.Range("A1").Value = "含交叉分析篩選器的樞紐分析表：可按部門與職級動態篩選薪資"
    With objPivotSheet.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
    End With

    ' ── 儲存 ────────────────────────────────────────────────────
    objWorkbook.SaveAs savePath, xlOpenXMLWorkbook

    Set objSlicerCache1 = Nothing
    Set objSlicerCache2 = Nothing
    Set objField        = Nothing
    Set objPivot        = Nothing
    Set objCache        = Nothing
    Set objPivotSheet   = Nothing
    Set objDataSheet    = Nothing
    Set objWorkbook     = Nothing

    MsgBox "完成！檔案已儲存至：" & savePath

End Sub
