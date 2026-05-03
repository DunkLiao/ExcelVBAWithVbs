Attribute VB_Name = "PivotWithRankNumber"
' ============================================================
' PivotWithRankNumber.bas
' 說明：使用 Excel VBA 自動建立顯示名次排名的樞紐分析表
' 功能：
'   1. 建立新活頁簿
'   2. 在「業務業績」工作表填入業務員銷售示範資料
'   3. 建立樞紐分析表（列=業務員）
'   4. 同時顯示業績金額及全體排名（降冪）
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 PivotWithRankNumber 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  As String = "業務業績"
Const SHEET_PIVOT As String = "樞紐分析表"
Const PIVOT_NAME  As String = "排名樞紐"
Const OUTPUT_FILE As String = "19_PivotWithRankNumber.xlsx"

Sub PivotWithRankNumber()

    ' ── 範例資料（部門、業務員、業績金額）──────────────────────
    ' 4 組 × 5 業務員 = 20 筆（每人 1 季業績）
    Dim arrDepts(19)   As String
    Dim arrEmps(19)    As String
    Dim arrAmounts(19) As Long

    arrDepts(0)  = "業務A組" : arrEmps(0)  = "王小明" : arrAmounts(0)  = 850000
    arrDepts(1)  = "業務A組" : arrEmps(1)  = "李大華" : arrAmounts(1)  = 720000
    arrDepts(2)  = "業務A組" : arrEmps(2)  = "陳美玲" : arrAmounts(2)  = 980000
    arrDepts(3)  = "業務A組" : arrEmps(3)  = "張志強" : arrAmounts(3)  = 650000
    arrDepts(4)  = "業務A組" : arrEmps(4)  = "林佳慧" : arrAmounts(4)  = 810000
    arrDepts(5)  = "業務B組" : arrEmps(5)  = "黃文成" : arrAmounts(5)  = 920000
    arrDepts(6)  = "業務B組" : arrEmps(6)  = "吳雅婷" : arrAmounts(6)  = 540000
    arrDepts(7)  = "業務B組" : arrEmps(7)  = "劉建宏" : arrAmounts(7)  = 1100000
    arrDepts(8)  = "業務B組" : arrEmps(8)  = "蔡玉芬" : arrAmounts(8)  = 780000
    arrDepts(9)  = "業務B組" : arrEmps(9)  = "謝俊傑" : arrAmounts(9)  = 690000
    arrDepts(10) = "業務C組" : arrEmps(10) = "鄭麗華" : arrAmounts(10) = 860000
    arrDepts(11) = "業務C組" : arrEmps(11) = "洪志明" : arrAmounts(11) = 750000
    arrDepts(12) = "業務C組" : arrEmps(12) = "許淑芬" : arrAmounts(12) = 620000
    arrDepts(13) = "業務C組" : arrEmps(13) = "楊建國" : arrAmounts(13) = 1050000
    arrDepts(14) = "業務C組" : arrEmps(14) = "廖秀琴" : arrAmounts(14) = 590000
    arrDepts(15) = "業務D組" : arrEmps(15) = "詹文彬" : arrAmounts(15) = 940000
    arrDepts(16) = "業務D組" : arrEmps(16) = "施淑娟" : arrAmounts(16) = 830000
    arrDepts(17) = "業務D組" : arrEmps(17) = "賴俊宏" : arrAmounts(17) = 710000
    arrDepts(18) = "業務D組" : arrEmps(18) = "江美惠" : arrAmounts(18) = 680000
    arrDepts(19) = "業務D組" : arrEmps(19) = "周志豪" : arrAmounts(19) = 1200000

    ' ── 主程式 ──────────────────────────────────────────────────
    Dim objWorkbook   As Workbook
    Dim objDataSheet  As Worksheet
    Dim objPivotSheet As Worksheet
    Dim objCache      As PivotCache
    Dim objPivot      As PivotTable
    Dim objField      As PivotField
    Dim objDataField  As PivotField
    Dim savePath      As String
    Dim i             As Integer

    savePath = Environ("USERPROFILE") & "\Desktop\" & OUTPUT_FILE

    Set objWorkbook   = Workbooks.Add()
    Set objDataSheet  = objWorkbook.Sheets(1)
    objDataSheet.Name = SHEET_DATA

    ' ── 寫入標題列 ──────────────────────────────────────────────
    objDataSheet.Cells(1, 1).Value = "部門"
    objDataSheet.Cells(1, 2).Value = "業務員"
    objDataSheet.Cells(1, 3).Value = "業績金額"

    With objDataSheet.Range("A1:C1")
        .Font.Bold           = True
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 19
        objDataSheet.Cells(i + 2, 1).Value = arrDepts(i)
        objDataSheet.Cells(i + 2, 2).Value = arrEmps(i)
        objDataSheet.Cells(i + 2, 3).Value = arrAmounts(i)
    Next i

    objDataSheet.Columns("A:C").AutoFit

    ' ── 新增樞紐分析表工作表 ─────────────────────────────────────
    Set objPivotSheet  = objWorkbook.Sheets.Add()
    objPivotSheet.Name = SHEET_PIVOT
    objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

    ' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
    Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:C21"))
    Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

    ' ── 設定列欄位 ──────────────────────────────────────────────
    Set objField = objPivot.PivotFields("業務員")
    objField.Orientation = xlRowField
    objField.Position    = 1

    ' ── 第一個值欄位：業績金額 ──────────────────────────────────
    Set objField = objPivot.PivotFields("業績金額")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "業績金額"

    ' ── 第二個值欄位：全體排名（降冪）────────────────────────────
    Set objField = objPivot.PivotFields("業績金額")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "全體排名"

    ' ── 設定排名計算方式（降冪，以業務員為基準欄位）────────────
    Set objDataField = objPivot.DataFields("全體排名")
    objDataField.Calculation = xlRankDecreasing
    objDataField.BaseField   = "業務員"

    ' ── 加入說明標題 ─────────────────────────────────────────────
    objPivotSheet.Range("A1").Value = "排名樞紐分析表：業務員業績金額及全體排名（降冪）"
    With objPivotSheet.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
    End With

    ' ── 儲存 ────────────────────────────────────────────────────
    objWorkbook.SaveAs savePath, xlOpenXMLWorkbook

    Set objDataField  = Nothing
    Set objField      = Nothing
    Set objPivot      = Nothing
    Set objCache      = Nothing
    Set objPivotSheet = Nothing
    Set objDataSheet  = Nothing
    Set objWorkbook   = Nothing

    MsgBox "完成！檔案已儲存至：" & savePath

End Sub
