Attribute VB_Name = "FilterAndSummarize"
Option Explicit

'************************************************************************************
' 模組名稱: FilterAndSummarize
' 功能說明: 篩選後使用 SUBTOTAL 函數對可見儲存格加總，並輸出分類彙總表
'           同時示範 SpecialCells(xlCellTypeVisible) 取得篩選後可見範圍
'
' 作者版權: Dunk
' 現任設計: Dunk
' 最後修改: 2026/5/9
'************************************************************************************

' 入口：對各區域套用篩選並產生彙總結果
Public Sub FilterAndSummarizeByRegion()
    On Error GoTo ErrHandler

    Dim wsData   As Worksheet
    Dim wsResult As Worksheet
    Dim regions  As Variant
    Dim i        As Integer

    Set wsData = GetOrCreateWsSum(ThisWorkbook, "銷售資料")
    Call FillSalesRegionData(wsData)

    Set wsResult = GetOrCreateWsSum(ThisWorkbook, "區域彙總")
    wsResult.Range("A1:C1").Value = Array("區域", "篩選筆數", "合計銷售額")

    regions = Array("北區", "中區", "南區", "東區")
    Dim resultRow As Integer
    resultRow = 2

    For i = 0 To UBound(regions)
        Dim regionName  As String
        Dim filteredCnt As Long
        Dim filteredSum As Double

        regionName = regions(i)
        Call GetFilteredStats(wsData, regionName, filteredCnt, filteredSum)

        wsResult.Cells(resultRow, 1).Value = regionName
        wsResult.Cells(resultRow, 2).Value = filteredCnt
        wsResult.Cells(resultRow, 3).Value = filteredSum
        resultRow = resultRow + 1
    Next i

    ' 清除篩選狀態
    If wsData.AutoFilterMode Then wsData.AutoFilterMode = False

    wsResult.Range("C2:C5").NumberFormat = "#,##0"
    wsResult.Columns("A:C").AutoFit
    wsResult.Activate
    MsgBox "各區域彙總完成，結果在「區域彙總」工作表。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    If wsData.AutoFilterMode Then wsData.AutoFilterMode = False
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 對指定區域篩選，取得可見列數與合計
Private Sub GetFilteredStats(ByVal ws As Worksheet, ByVal region As String, _
                              ByRef outCount As Long, ByRef outSum As Double)
    Dim rng         As Range
    Dim visibleRng  As Range
    Dim lastRow     As Long

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    Set rng = ws.Range("A1").CurrentRegion
    rng.AutoFilter Field:=2, Criteria1:=region

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    On Error Resume Next
    Set visibleRng = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If visibleRng Is Nothing Then
        outCount = 0
        outSum = 0
    Else
        outCount = visibleRng.Count
        ' SUBTOTAL(9,...) 僅加總可見儲存格
        outSum = Application.WorksheetFunction.Subtotal(9, ws.Range("C2:C" & lastRow))
    End If
End Sub

' 填入多區域銷售測試資料
Private Sub FillSalesRegionData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("月份", "區域", "銷售額")
    ws.Range("A2:C2").Value = Array("2026/01", "北區", 580000)
    ws.Range("A3:C3").Value = Array("2026/01", "中區", 420000)
    ws.Range("A4:C4").Value = Array("2026/01", "南區", 650000)
    ws.Range("A5:C5").Value = Array("2026/01", "東區", 310000)
    ws.Range("A6:C6").Value = Array("2026/02", "北區", 720000)
    ws.Range("A7:C7").Value = Array("2026/02", "中區", 390000)
    ws.Range("A8:C8").Value = Array("2026/02", "南區", 480000)
    ws.Range("A9:C9").Value = Array("2026/02", "東區", 260000)
    ws.Range("C2:C9").NumberFormat = "#,##0"
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表並清空
Private Function GetOrCreateWsSum(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateWsSum = ws
End Function