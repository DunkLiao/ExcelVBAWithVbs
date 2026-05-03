Attribute VB_Name = "PivotWithShowPages"
' ============================================================
' PivotWithShowPages.bas
' 說明：使用 Excel VBA 自動建立並展開「顯示報表篩選頁面」的樞紐分析表
' 功能：
'   1. 建立新活頁簿
'   2. 在「銷售資料」工作表填入多地區季度銷售示範資料
'   3. 建立含地區篩選頁的樞紐分析表
'   4. 呼叫 ShowPages 方法，自動依每個地區值建立獨立工作表
'   5. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 PivotWithShowPages 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  As String = "銷售資料"
Const SHEET_PIVOT As String = "樞紐分析表"
Const PIVOT_NAME  As String = "分頁展開樞紐"
Const OUTPUT_FILE As String = "17_PivotWithShowPages.xlsx"

Sub PivotWithShowPages()

    ' ── 範例資料（地區、季度、產品、銷售額）────────────────────
    ' 4 地區 × 3 季度 × 2 產品 = 24 筆
    Dim arrRegions(23)  As String
    Dim arrQtrs(23)     As String
    Dim arrProducts(23) As String
    Dim arrAmounts(23)  As Long

    arrRegions(0)  = "北部" : arrQtrs(0)  = "Q1" : arrProducts(0)  = "筆電" : arrAmounts(0)  = 185000
    arrRegions(1)  = "北部" : arrQtrs(1)  = "Q1" : arrProducts(1)  = "平板" : arrAmounts(1)  = 112000
    arrRegions(2)  = "北部" : arrQtrs(2)  = "Q2" : arrProducts(2)  = "筆電" : arrAmounts(2)  = 210000
    arrRegions(3)  = "北部" : arrQtrs(3)  = "Q2" : arrProducts(3)  = "平板" : arrAmounts(3)  = 135000
    arrRegions(4)  = "北部" : arrQtrs(4)  = "Q3" : arrProducts(4)  = "筆電" : arrAmounts(4)  = 245000
    arrRegions(5)  = "北部" : arrQtrs(5)  = "Q3" : arrProducts(5)  = "平板" : arrAmounts(5)  = 158000
    arrRegions(6)  = "中部" : arrQtrs(6)  = "Q1" : arrProducts(6)  = "筆電" : arrAmounts(6)  = 142000
    arrRegions(7)  = "中部" : arrQtrs(7)  = "Q1" : arrProducts(7)  = "平板" : arrAmounts(7)  = 88000
    arrRegions(8)  = "中部" : arrQtrs(8)  = "Q2" : arrProducts(8)  = "筆電" : arrAmounts(8)  = 165000
    arrRegions(9)  = "中部" : arrQtrs(9)  = "Q2" : arrProducts(9)  = "平板" : arrAmounts(9)  = 102000
    arrRegions(10) = "中部" : arrQtrs(10) = "Q3" : arrProducts(10) = "筆電" : arrAmounts(10) = 190000
    arrRegions(11) = "中部" : arrQtrs(11) = "Q3" : arrProducts(11) = "平板" : arrAmounts(11) = 118000
    arrRegions(12) = "南部" : arrQtrs(12) = "Q1" : arrProducts(12) = "筆電" : arrAmounts(12) = 158000
    arrRegions(13) = "南部" : arrQtrs(13) = "Q1" : arrProducts(13) = "平板" : arrAmounts(13) = 95000
    arrRegions(14) = "南部" : arrQtrs(14) = "Q2" : arrProducts(14) = "筆電" : arrAmounts(14) = 178000
    arrRegions(15) = "南部" : arrQtrs(15) = "Q2" : arrProducts(15) = "平板" : arrAmounts(15) = 108000
    arrRegions(16) = "南部" : arrQtrs(16) = "Q3" : arrProducts(16) = "筆電" : arrAmounts(16) = 205000
    arrRegions(17) = "南部" : arrQtrs(17) = "Q3" : arrProducts(17) = "平板" : arrAmounts(17) = 125000
    arrRegions(18) = "東部" : arrQtrs(18) = "Q1" : arrProducts(18) = "筆電" : arrAmounts(18) = 68000
    arrRegions(19) = "東部" : arrQtrs(19) = "Q1" : arrProducts(19) = "平板" : arrAmounts(19) = 42000
    arrRegions(20) = "東部" : arrQtrs(20) = "Q2" : arrProducts(20) = "筆電" : arrAmounts(20) = 75000
    arrRegions(21) = "東部" : arrQtrs(21) = "Q2" : arrProducts(21) = "平板" : arrAmounts(21) = 48000
    arrRegions(22) = "東部" : arrQtrs(22) = "Q3" : arrProducts(22) = "筆電" : arrAmounts(22) = 88000
    arrRegions(23) = "東部" : arrQtrs(23) = "Q3" : arrProducts(23) = "平板" : arrAmounts(23) = 55000

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
    objDataSheet.Cells(1, 1).Value = "地區"
    objDataSheet.Cells(1, 2).Value = "季度"
    objDataSheet.Cells(1, 3).Value = "產品"
    objDataSheet.Cells(1, 4).Value = "銷售額"

    With objDataSheet.Range("A1:D1")
        .Font.Bold           = True
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 23
        objDataSheet.Cells(i + 2, 1).Value = arrRegions(i)
        objDataSheet.Cells(i + 2, 2).Value = arrQtrs(i)
        objDataSheet.Cells(i + 2, 3).Value = arrProducts(i)
        objDataSheet.Cells(i + 2, 4).Value = arrAmounts(i)
    Next i

    objDataSheet.Columns("A:D").AutoFit

    ' ── 新增樞紐分析表工作表 ─────────────────────────────────────
    Set objPivotSheet  = objWorkbook.Sheets.Add()
    objPivotSheet.Name = SHEET_PIVOT
    objPivotSheet.Move , objWorkbook.Sheets(objWorkbook.Sheets.Count)

    ' ── 建立樞紐快取與樞紐分析表 ─────────────────────────────────
    Set objCache = objWorkbook.PivotCaches.Create(xlDatabase, objDataSheet.Range("A1:D25"))
    Set objPivot = objCache.CreatePivotTable(objPivotSheet.Range("A3"), PIVOT_NAME)

    ' ── 設定篩選、列、欄、值欄位 ─────────────────────────────────
    Set objField = objPivot.PivotFields("地區")
    objField.Orientation = xlPageField
    objField.Position    = 1

    Set objField = objPivot.PivotFields("季度")
    objField.Orientation = xlRowField
    objField.Position    = 1

    Set objField = objPivot.PivotFields("產品")
    objField.Orientation = xlColumnField
    objField.Position    = 1

    Set objField = objPivot.PivotFields("銷售額")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "加總 - 銷售額"

    ' ── 顯示報表篩選頁面：依「地區」自動建立各地區獨立工作表 ──
    ' ShowPages 會依篩選欄位的每個項目值，複製樞紐並建立對應工作表
    objPivot.ShowPages "地區"

    ' ── 加入說明標題 ─────────────────────────────────────────────
    objPivotSheet.Range("A1").Value = "顯示報表篩選頁面：ShowPages 自動依地區建立各地區獨立工作表"
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

    MsgBox "完成！" & vbCrLf & _
           "已依地區展開獨立工作表（北部/中部/南部/東部）" & vbCrLf & _
           "檔案已儲存至：" & savePath

End Sub
