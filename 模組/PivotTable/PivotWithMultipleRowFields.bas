Attribute VB_Name = "PivotWithMultipleRowFields"
' ============================================================
' PivotWithMultipleRowFields.bas
' 說明：使用 Excel VBA 自動建立含三層巢狀列欄位的樞紐分析表
' 功能：
'   1. 建立新活頁簿
'   2. 在「業績資料」工作表填入多維度銷售示範資料
'   3. 建立樞紐分析表（列=大區 > 城市 > 通路，三層巢狀）
'   4. 將結果儲存至桌面
' 執行方式：在 Excel VBA 編輯器中執行 PivotWithMultipleRowFields 巨集
' ============================================================

Option Explicit

' ── 常數設定 ────────────────────────────────────────────────
Const SHEET_DATA  As String = "業績資料"
Const SHEET_PIVOT As String = "樞紐分析表"
Const PIVOT_NAME  As String = "多層列欄位樞紐"
Const OUTPUT_FILE As String = "14_PivotWithMultipleRowFields.xlsx"

Sub PivotWithMultipleRowFields()

    ' ── 範例資料（大區、城市、通路、銷售額）────────────────────
    Dim arrZones(23)    As String
    Dim arrCities(23)   As String
    Dim arrChannels(23) As String
    Dim arrAmounts(23)  As Long

    arrZones(0)  = "北部" : arrCities(0)  = "台北" : arrChannels(0)  = "直營" : arrAmounts(0)  = 210000
    arrZones(1)  = "北部" : arrCities(1)  = "台北" : arrChannels(1)  = "代理" : arrAmounts(1)  = 145000
    arrZones(2)  = "北部" : arrCities(2)  = "基隆" : arrChannels(2)  = "直營" : arrAmounts(2)  = 68000
    arrZones(3)  = "北部" : arrCities(3)  = "基隆" : arrChannels(3)  = "代理" : arrAmounts(3)  = 52000
    arrZones(4)  = "北部" : arrCities(4)  = "桃園" : arrChannels(4)  = "直營" : arrAmounts(4)  = 175000
    arrZones(5)  = "北部" : arrCities(5)  = "桃園" : arrChannels(5)  = "代理" : arrAmounts(5)  = 130000
    arrZones(6)  = "中部" : arrCities(6)  = "台中" : arrChannels(6)  = "直營" : arrAmounts(6)  = 195000
    arrZones(7)  = "中部" : arrCities(7)  = "台中" : arrChannels(7)  = "代理" : arrAmounts(7)  = 160000
    arrZones(8)  = "中部" : arrCities(8)  = "彰化" : arrChannels(8)  = "直營" : arrAmounts(8)  = 78000
    arrZones(9)  = "中部" : arrCities(9)  = "彰化" : arrChannels(9)  = "代理" : arrAmounts(9)  = 61000
    arrZones(10) = "南部" : arrCities(10) = "台南" : arrChannels(10) = "直營" : arrAmounts(10) = 152000
    arrZones(11) = "南部" : arrCities(11) = "台南" : arrChannels(11) = "代理" : arrAmounts(11) = 118000
    arrZones(12) = "南部" : arrCities(12) = "高雄" : arrChannels(12) = "直營" : arrAmounts(12) = 231000
    arrZones(13) = "南部" : arrCities(13) = "高雄" : arrChannels(13) = "代理" : arrAmounts(13) = 187000
    arrZones(14) = "南部" : arrCities(14) = "屏東" : arrChannels(14) = "直營" : arrAmounts(14) = 58000
    arrZones(15) = "南部" : arrCities(15) = "屏東" : arrChannels(15) = "代理" : arrAmounts(15) = 43000
    arrZones(16) = "東部" : arrCities(16) = "花蓮" : arrChannels(16) = "直營" : arrAmounts(16) = 47000
    arrZones(17) = "東部" : arrCities(17) = "花蓮" : arrChannels(17) = "代理" : arrAmounts(17) = 35000
    arrZones(18) = "東部" : arrCities(18) = "台東" : arrChannels(18) = "直營" : arrAmounts(18) = 39000
    arrZones(19) = "東部" : arrCities(19) = "台東" : arrChannels(19) = "代理" : arrAmounts(19) = 28000
    arrZones(20) = "離島" : arrCities(20) = "澎湖" : arrChannels(20) = "直營" : arrAmounts(20) = 32000
    arrZones(21) = "離島" : arrCities(21) = "澎湖" : arrChannels(21) = "代理" : arrAmounts(21) = 21000
    arrZones(22) = "離島" : arrCities(22) = "金門" : arrChannels(22) = "直營" : arrAmounts(22) = 25000
    arrZones(23) = "離島" : arrCities(23) = "金門" : arrChannels(23) = "代理" : arrAmounts(23) = 17000

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
    objDataSheet.Cells(1, 1).Value = "大區"
    objDataSheet.Cells(1, 2).Value = "城市"
    objDataSheet.Cells(1, 3).Value = "通路"
    objDataSheet.Cells(1, 4).Value = "銷售額"

    With objDataSheet.Range("A1:D1")
        .Font.Bold           = True
        .Interior.Color      = RGB(68, 114, 196)
        .Font.Color          = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' ── 寫入資料列 ──────────────────────────────────────────────
    For i = 0 To 23
        objDataSheet.Cells(i + 2, 1).Value = arrZones(i)
        objDataSheet.Cells(i + 2, 2).Value = arrCities(i)
        objDataSheet.Cells(i + 2, 3).Value = arrChannels(i)
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

    ' ── 設定三層巢狀列欄位 ──────────────────────────────────────
    Set objField = objPivot.PivotFields("大區")
    objField.Orientation = xlRowField
    objField.Position    = 1

    Set objField = objPivot.PivotFields("城市")
    objField.Orientation = xlRowField
    objField.Position    = 2

    Set objField = objPivot.PivotFields("通路")
    objField.Orientation = xlRowField
    objField.Position    = 3

    ' ── 設定值欄位 ──────────────────────────────────────────────
    Set objField = objPivot.PivotFields("銷售額")
    objField.Orientation = xlDataField
    objField.Function    = xlSum
    objField.Name        = "加總 - 銷售額"

    ' ── 加入說明標題 ─────────────────────────────────────────────
    objPivotSheet.Range("A1").Value = "多層列欄位樞紐分析表：大區 > 城市 > 通路 三層巢狀列欄位"
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
