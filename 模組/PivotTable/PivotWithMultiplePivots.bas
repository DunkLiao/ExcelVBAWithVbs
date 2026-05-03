Attribute VB_Name = "PivotWithMultiplePivots"
' ============================================================
' PivotWithMultiplePivots.bas
' 說明：使用 Excel VBA 自動建立共用快取的兩個樞紐分析表
' 功能：
'   1. 建立新活頁簿
'   2. 在「招募資料」工作表填入各部門招募示範資料
'   3. 以同一個樞紐快取建立兩個樞紐分析表，共用同一份資料來源
'   4. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 PivotWithMultiplePivots 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  As String = "招募資料"
Const SHEET_PIVOT As String = "樞紐分析表"
Const OUTPUT_FILE As String = "18_PivotWithMultiplePivots.xlsx"

Sub PivotWithMultiplePivots()

    ' ── 範例資料（部門、招募管道、招募人數、招募費用）──────────
    Dim arrDepts(19)    As String
    Dim arrChannels(19) As String
    Dim arrCounts(19)   As Long
    Dim arrCosts(19)    As Long

    arrDepts(0)  = "研發部" : arrChannels(0)  = "獵頭公司" : arrCounts(0)  = 3  : arrCosts(0)  = 150000
    arrDepts(1)  = "研發部" : arrChannels(1)  = "104人力銀行" : arrCounts(1)  = 5  : arrCosts(1)  = 30000
    arrDepts(2)  = "研發部" : arrChannels(2)  = "校園徵才" : arrCounts(2)  = 2  : arrCosts(2)  = 20000
    arrDepts(3)  = "研發部" : arrChannels(3)  = "員工推薦" : arrCounts(3)  = 1  : arrCosts(3)  = 10000
    arrDepts(4)  = "研發部" : arrChannels(4)  = "LinkedIn" : arrCounts(4)  = 2  : arrCosts(4)  = 25000
    arrDepts(5)  = "業務部" : arrChannels(5)  = "獵頭公司" : arrCounts(5)  = 2  : arrCosts(5)  = 80000
    arrDepts(6)  = "業務部" : arrChannels(6)  = "104人力銀行" : arrCounts(6)  = 8  : arrCosts(6)  = 48000
    arrDepts(7)  = "業務部" : arrChannels(7)  = "校園徵才" : arrCounts(7)  = 4  : arrCosts(7)  = 40000
    arrDepts(8)  = "業務部" : arrChannels(8)  = "員工推薦" : arrCounts(8)  = 3  : arrCosts(8)  = 30000
    arrDepts(9)  = "業務部" : arrChannels(9)  = "LinkedIn" : arrCounts(9)  = 1  : arrCosts(9)  = 12000
    arrDepts(10) = "生產部" : arrChannels(10) = "獵頭公司" : arrCounts(10) = 1  : arrCosts(10) = 40000
    arrDepts(11) = "生產部" : arrChannels(11) = "104人力銀行" : arrCounts(11) = 12 : arrCosts(11) = 60000
    arrDepts(12) = "生產部" : arrChannels(12) = "校園徵才" : arrCounts(12) = 6  : arrCosts(12) = 50000
    arrDepts(13) = "生產部" : arrChannels(13) = "員工推薦" : arrCounts(13) = 5  : arrCosts(13) = 50000
    arrDepts(14) = "生產部" : arrChannels(14) = "LinkedIn" : arrCounts(14) = 0  : arrCosts(14) = 0
    arrDepts(15) = "行政部" : arrChannels(15) = "獵頭公司" : arrCounts(15) = 0  : arrCosts(15) = 0
    arrDepts(16) = "行政部" : arrChannels(16) = "104人力銀行" : arrCounts(16) = 4  : arrCosts(16) = 24000
    arrDepts(17) = "行政部" : arrChannels(17) = "校園徵才" : arrCounts(17) = 1  : arrCosts(17) = 10000
    arrDepts(18) = "行政部" : arrChannels(18) = "員工推薦" : arrCounts(18) = 2  : arrCosts(18) = 20000
    arrDepts(19) = "行政部" : arrChannels(19) = "LinkedIn" : arrCounts(19) = 1  : arrCosts(19) = 12000

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook   As Workbook
    Dim objDataSheet  As Worksheet
    Dim objPivotSheet As Worksheet
    Dim objCache      As PivotCache
    Dim objPivot1     As PivotTable
    Dim objPivot2     As PivotTable
    Dim objField      As PivotField
    Dim savePath      As String
    Dim i             As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook   = Workbooks.Add()
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
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 19
        objDataSheet.Cells(i + 2, 1).Value = arrDepts(i)
        objDataSheet.Cells(i + 2, 2).Value = arrChannels(i)
        objDataSheet.Cells(i + 2, 3).Value = arrCounts(i)
        objDataSheet.Cells(i + 2, 4).Value = arrCosts(i)
    Next i

    objDataSheet.Columns("A:D").AutoFit

    ' ── 新增樞紐分析表工作表 ─────────────────────────────────────
    Set objPivotSheet  = objWorkbook.Sheets.Add()
    objPivotSheet.Name = SHEET_PIVOT
    objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

    ' ── 建立共用樞紐快取 ─────────────────────────────────────────
    Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:D21"))

    ' ── 樞紐一：部門招募人數（A3）────────────────────────────────
    Set objPivot1 = objCache.CreatePivotTable(objPivotSheet.Range("A3"), "部門招募人數")

    Set objField = objPivot1.PivotFields("部門")
    objField.Orientation = xlRowField
    objField.Position    = 1

    Set objField = objPivot1.PivotFields("招募人數")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "加總 - 招募人數"

    ' ── 樞紐二：管道費用分析（H3），與樞紐一共用快取 ────────────
    Set objPivot2 = objCache.CreatePivotTable(objPivotSheet.Range("H3"), "管道費用分析")

    Set objField = objPivot2.PivotFields("招募管道")
    objField.Orientation = xlRowField
    objField.Position    = 1

    Set objField = objPivot2.PivotFields("招募費用")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "加總 - 招募費用"

    ' ── 加入說明標題 ─────────────────────────────────────────────
    objPivotSheet.Range("A1").Value = "多樞紐共用快取：樞紐一（部門招募人數）與樞紐二（管道費用）共享同一資料快取"
    With objPivotSheet.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
    End With

    ' ── 儲存 ────────────────────────────────────────────────────
    objWorkbook.SaveAs savePath, xlOpenXMLWorkbook

    Set objField      = Nothing
    Set objPivot1     = Nothing
    Set objPivot2     = Nothing
    Set objCache      = Nothing
    Set objPivotSheet = Nothing
    Set objDataSheet  = Nothing
    Set objWorkbook   = Nothing

    MsgBox "完成！檔案已儲存至：" & savePath

End Sub
