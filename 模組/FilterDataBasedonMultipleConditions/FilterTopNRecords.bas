Attribute VB_Name = "FilterTopNRecords"
Option Explicit

'************************************************************************************
' 模組名稱: FilterTopNRecords
' 功能說明: 使用 AutoFilter xlTop10Items / xlTop10Percent 篩選前 N 筆或前 N%
'           示範依業績欄位篩選頂尖業務員
'
' 作者版權: Dunk
' 現任設計: Dunk
' 最後修改: 2026/5/9
'************************************************************************************

' 入口：篩選業績前 3 名
Public Sub FilterTop3SalesExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateWsClean(ThisWorkbook, "前N名篩選範例")
    Call FillSalesData(ws)
    Call ApplyTopNFilter(ws, 3, False)

    MsgBox "已篩選出業績前 3 名業務員。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 入口：篩選業績前 20%
Public Sub FilterTop20PercentSalesExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateWsClean(ThisWorkbook, "前20%篩選範例")
    Call FillSalesData(ws)
    Call ApplyTopNFilter(ws, 20, True)

    MsgBox "已篩選出業績前 20% 業務員。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 套用前 N 名或前 N% 篩選
' isPercent=True 代表百分比，False 代表固定筆數
Private Sub ApplyTopNFilter(ByVal ws As Worksheet, _
                             ByVal topN As Integer, _
                             ByVal isPercent As Boolean)
    Dim rng As Range
    Dim op As XlAutoFilterOperator

    Set rng = ws.Range("A1").CurrentRegion
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    If isPercent Then
        op = xlTop10Percent
    Else
        op = xlTop10Items
    End If

    ' Field:=3 業績欄（C欄）取前 topN
    rng.AutoFilter Field:=3, Criteria1:=topN, Operator:=op
    ws.Columns("A:D").AutoFit
End Sub

' 填入業務測試資料
Private Sub FillSalesData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("業務員", "區域", "業績(元)", "達成率(%)")
    ws.Range("A2:D2").Value = Array("王小明", "北區", 980000, 112)
    ws.Range("A3:D3").Value = Array("李美玲", "中區", 760000, 95)
    ws.Range("A4:D4").Value = Array("張志豪", "南區", 1250000, 130)
    ws.Range("A5:D5").Value = Array("陳雅婷", "東區", 540000, 78)
    ws.Range("A6:D6").Value = Array("林建宏", "北區", 870000, 105)
    ws.Range("A7:D7").Value = Array("吳惠君", "中區", 430000, 65)
    ws.Range("A8:D8").Value = Array("黃志偉", "南區", 1100000, 120)
    ws.Range("A9:D9").Value = Array("蔡佳蓉", "東區", 690000, 88)
    ws.Range("C2:C9").NumberFormat = "#,##0"
    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表並清空
Private Function GetOrCreateWsClean(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateWsClean = ws
End Function