Attribute VB_Name = "PivotWithTopN"
' ============================================================
' PivotWithTopN.bas
' 說明：使用 Excel VBA 自動建立以 AutoShow 篩選前 N 名的樞紐分析表
' 功能：
'   1. 建立新活頁簿
'   2. 在「業績資料」工作表填入業務員季度銷售示範資料
'   3. 建立樞紐分析表（列=業務員，值=銷售額加總）
'   4. 使用 AutoShow 只顯示銷售額最高的前 N 名業務員
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 PivotWithTopN 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  As String = "業績資料"
Const SHEET_PIVOT As String = "樞紐分析表"
Const PIVOT_NAME  As String = "前N名樞紐"
Const OUTPUT_FILE As String = "07_PivotWithTopN.xlsx"
Const TOP_N       As Long   = 5

Sub PivotWithTopN()

    ' ── 範例資料（業務員、季度、銷售額）────────────────────────
    ' 7 業務員 × 4 季度 = 28 筆
    Dim arrEmps(27)    As String
    Dim arrQtrs(27)    As String
    Dim arrAmounts(27) As Long

    arrEmps(0)  = "王小明" : arrQtrs(0)  = "Q1" : arrAmounts(0)  = 215000
    arrEmps(1)  = "王小明" : arrQtrs(1)  = "Q2" : arrAmounts(1)  = 245000
    arrEmps(2)  = "王小明" : arrQtrs(2)  = "Q3" : arrAmounts(2)  = 268000
    arrEmps(3)  = "王小明" : arrQtrs(3)  = "Q4" : arrAmounts(3)  = 310000
    arrEmps(4)  = "李大華" : arrQtrs(4)  = "Q1" : arrAmounts(4)  = 185000
    arrEmps(5)  = "李大華" : arrQtrs(5)  = "Q2" : arrAmounts(5)  = 192000
    arrEmps(6)  = "李大華" : arrQtrs(6)  = "Q3" : arrAmounts(6)  = 178000
    arrEmps(7)  = "李大華" : arrQtrs(7)  = "Q4" : arrAmounts(7)  = 220000
    arrEmps(8)  = "陳美玲" : arrQtrs(8)  = "Q1" : arrAmounts(8)  = 305000
    arrEmps(9)  = "陳美玲" : arrQtrs(9)  = "Q2" : arrAmounts(9)  = 328000
    arrEmps(10) = "陳美玲" : arrQtrs(10) = "Q3" : arrAmounts(10) = 351000
    arrEmps(11) = "陳美玲" : arrQtrs(11) = "Q4" : arrAmounts(11) = 395000
    arrEmps(12) = "張志強" : arrQtrs(12) = "Q1" : arrAmounts(12) = 158000
    arrEmps(13) = "張志強" : arrQtrs(13) = "Q2" : arrAmounts(13) = 165000
    arrEmps(14) = "張志強" : arrQtrs(14) = "Q3" : arrAmounts(14) = 142000
    arrEmps(15) = "張志強" : arrQtrs(15) = "Q4" : arrAmounts(15) = 180000
    arrEmps(16) = "林佳慧" : arrQtrs(16) = "Q1" : arrAmounts(16) = 275000
    arrEmps(17) = "林佳慧" : arrQtrs(17) = "Q2" : arrAmounts(17) = 290000
    arrEmps(18) = "林佳慧" : arrQtrs(18) = "Q3" : arrAmounts(18) = 315000
    arrEmps(19) = "林佳慧" : arrQtrs(19) = "Q4" : arrAmounts(19) = 358000
    arrEmps(20) = "黃文成" : arrQtrs(20) = "Q1" : arrAmounts(20) = 128000
    arrEmps(21) = "黃文成" : arrQtrs(21) = "Q2" : arrAmounts(21) = 135000
    arrEmps(22) = "黃文成" : arrQtrs(22) = "Q3" : arrAmounts(22) = 118000
    arrEmps(23) = "黃文成" : arrQtrs(23) = "Q4" : arrAmounts(23) = 148000
    arrEmps(24) = "吳雅婷" : arrQtrs(24) = "Q1" : arrAmounts(24) = 238000
    arrEmps(25) = "吳雅婷" : arrQtrs(25) = "Q2" : arrAmounts(25) = 252000
    arrEmps(26) = "吳雅婷" : arrQtrs(26) = "Q3" : arrAmounts(26) = 275000
    arrEmps(27) = "吳雅婷" : arrQtrs(27) = "Q4" : arrAmounts(27) = 315000

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook   As Workbook
    Dim objDataSheet  As Worksheet
    Dim objPivotSheet As Worksheet
    Dim objCache      As PivotCache
    Dim objPivot      As PivotTable
    Dim objField      As PivotField
    Dim savePath      As String
    Dim i             As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook   = Workbooks.Add()
    Set objDataSheet  = objWorkbook.Sheets(1)
    objDataSheet.Name = SHEET_DATA

    ' ── 寫入標題列 ──────────────────────────────────────────────
    objDataSheet.Cells(1, 1).Value = "業務員"
    objDataSheet.Cells(1, 2).Value = "季度"
    objDataSheet.Cells(1, 3).Value = "銷售額"

    With objDataSheet.Range("A1:C1")
        .Font.Bold           = True
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 27
        objDataSheet.Cells(i + 2, 1).Value = arrEmps(i)
        objDataSheet.Cells(i + 2, 2).Value = arrQtrs(i)
        objDataSheet.Cells(i + 2, 3).Value = arrAmounts(i)
    Next i

    objDataSheet.Columns("A:C").AutoFit

    ' ── 新增樞紐分析表工作表 ─────────────────────────────────────
    Set objPivotSheet  = objWorkbook.Sheets.Add()
    objPivotSheet.Name = SHEET_PIVOT
    objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

    ' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
    Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C29"))
    Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

    ' ── 設定列、值欄位 ──────────────────────────────────────────
    Set objField = objPivot.PivotFields("業務員")
    objField.Orientation = xlRowField
    objField.Position    = 1

    Set objField = objPivot.PivotFields("銷售額")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "加總 - 銷售額"

    ' ── 篩選前 N 名：AutoShow(xlAutomatic, xlTop, N, 欄位名稱) ──
    ' 只顯示「加總 - 銷售額」最高的前 TOP_N 名業務員
    objPivot.PivotFields("業務員").AutoShow xlAutomatic, xlTop, TOP_N, "加總 - 銷售額"

    ' ── 加入說明標題 ─────────────────────────────────────────────
    objPivotSheet.Range("A1").Value = "前 N 名篩選樞紐分析表：銷售額前 " & TOP_N & " 名業務員"
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
