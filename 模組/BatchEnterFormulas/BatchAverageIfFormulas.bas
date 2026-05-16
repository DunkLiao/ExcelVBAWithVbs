Attribute VB_Name = "BatchAverageIfFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchAverageIfFormulas
'功能說明: 批次在指定欄位輸入 AVERAGEIF / AVERAGEIFS 公式，
'          依條件欄與求平均欄自動產生條件平均公式並填入結果區域
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

' 主程式：建立範例資料並批次輸入 AVERAGEIF 公式
Sub BatchInsertAverageIfFormulas()
    Dim ws As Worksheet
    Set ws = GetOrCreateAvgSheet("條件平均範例")
    ws.Cells.Clear

    ' 建立原始資料（A~D欄）
    Call FillAverageIfSampleData(ws)

    ' 批次輸入 AVERAGEIF 公式（F欄：依地區求銷售量平均）
    Call BuildAverageIfByRegion(ws)

    ws.Columns.AutoFit
    MsgBox "AVERAGEIF 公式已批次輸入完成！", vbInformation, "完成"
End Sub

' 批次輸入 AVERAGEIFS 公式（多條件）
Sub BatchInsertAverageIfsFormulas()
    Dim ws As Worksheet
    Set ws = GetOrCreateAvgSheet("多條件平均範例")
    ws.Cells.Clear

    Call FillAverageIfSampleData(ws)
    Call BuildAverageIfsMultiCondition(ws)

    ws.Columns.AutoFit
    MsgBox "AVERAGEIFS 公式已批次輸入完成！", vbInformation, "完成"
End Sub

' 依地區批次建立 AVERAGEIF 公式
Private Sub BuildAverageIfByRegion(ByVal ws As Worksheet)
    ' 匯總區域標題（F欄）
    ws.Range("F1").Value = "地區"
    ws.Range("G1").Value = "平均銷售量(AVERAGEIF)"

    Dim regions(1 To 3) As String
    regions(1) = "北區"
    regions(2) = "中區"
    regions(3) = "南區"

    Dim r As Integer
    For r = 1 To 3
        ws.Cells(r + 1, 6).Value = regions(r)
        ' AVERAGEIF(條件範圍, 條件值, 平均值範圍)
        ' 條件欄 = B（地區），平均值欄 = D（銷售量）
        ws.Cells(r + 1, 7).Formula = _
            "=AVERAGEIF($B$2:$B$13,F" & (r + 1) & ",$D$2:$D$13)"
        ws.Cells(r + 1, 7).NumberFormat = "#,##0.0"
    Next r

    ' 標題格式
    ws.Range("F1:G1").Font.Bold = True
    ws.Range("F1:G1").Interior.Color = RGB(198, 224, 180)
End Sub

' 建立多條件 AVERAGEIFS 公式
Private Sub BuildAverageIfsMultiCondition(ByVal ws As Worksheet)
    ws.Range("F1").Value = "地區"
    ws.Range("G1").Value = "產品"
    ws.Range("H1").Value = "平均銷售量(AVERAGEIFS)"

    ' 條件組合
    Dim arrRegion(1 To 4)  As String
    Dim arrProduct(1 To 4) As String
    arrRegion(1) = "北區": arrProduct(1) = "A產品"
    arrRegion(2) = "北區": arrProduct(2) = "B產品"
    arrRegion(3) = "南區": arrProduct(3) = "A產品"
    arrRegion(4) = "南區": arrProduct(4) = "B產品"

    Dim r As Integer
    For r = 1 To 4
        ws.Cells(r + 1, 6).Value = arrRegion(r)
        ws.Cells(r + 1, 7).Value = arrProduct(r)
        ' AVERAGEIFS(平均值範圍, 條件範圍1, 條件1, 條件範圍2, 條件2)
        ws.Cells(r + 1, 8).Formula = _
            "=AVERAGEIFS($D$2:$D$13,$B$2:$B$13,F" & (r + 1) & _
            ",$C$2:$C$13,G" & (r + 1) & ")"
        ws.Cells(r + 1, 8).NumberFormat = "#,##0.0"
    Next r

    ws.Range("F1:H1").Font.Bold = True
    ws.Range("F1:H1").Interior.Color = RGB(198, 224, 180)
End Sub

Private Sub FillAverageIfSampleData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("日期", "地區", "產品", "銷售量")
    ws.Range("A1:D1").Font.Bold = True

    Dim data(1 To 12, 1 To 4) As Variant
    data(1, 1) = "2026/1/5":  data(1, 2) = "北區": data(1, 3) = "A產品": data(1, 4) = 230
    data(2, 1) = "2026/1/8":  data(2, 2) = "中區": data(2, 3) = "B產品": data(2, 4) = 185
    data(3, 1) = "2026/1/12": data(3, 2) = "南區": data(3, 3) = "A產品": data(3, 4) = 310
    data(4, 1) = "2026/2/3":  data(4, 2) = "北區": data(4, 3) = "B產品": data(4, 4) = 265
    data(5, 1) = "2026/2/15": data(5, 2) = "中區": data(5, 3) = "A產品": data(5, 4) = 200
    data(6, 1) = "2026/2/20": data(6, 2) = "南區": data(6, 3) = "B產品": data(6, 4) = 175
    data(7, 1) = "2026/3/4":  data(7, 2) = "北區": data(7, 3) = "A產品": data(7, 4) = 290
    data(8, 1) = "2026/3/9":  data(8, 2) = "中區": data(8, 3) = "B產品": data(8, 4) = 220
    data(9, 1) = "2026/3/18": data(9, 2) = "南區": data(9, 3) = "A產品": data(9, 4) = 340
    data(10, 1) = "2026/4/2": data(10, 2) = "北區": data(10, 3) = "B產品": data(10, 4) = 195
    data(11, 1) = "2026/4/11":data(11, 2) = "中區": data(11, 3) = "A產品": data(11, 4) = 250
    data(12, 1) = "2026/4/22":data(12, 2) = "南區": data(12, 3) = "B產品": data(12, 4) = 280

    Dim r As Integer
    For r = 1 To 12
        ws.Cells(r + 1, 1).Value = CDate(data(r, 1))
        ws.Cells(r + 1, 1).NumberFormat = "yyyy/mm/dd"
        ws.Cells(r + 1, 2).Value = data(r, 2)
        ws.Cells(r + 1, 3).Value = data(r, 3)
        ws.Cells(r + 1, 4).Value = data(r, 4)
    Next r
End Sub

Private Function GetOrCreateAvgSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateAvgSheet = ws
End Function
