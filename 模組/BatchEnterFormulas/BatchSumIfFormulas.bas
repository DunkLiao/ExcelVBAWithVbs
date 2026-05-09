Attribute VB_Name = "BatchSumIfFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchSumIfFormulas
'功能說明: 批次填入 SUMIF / SUMIFS 條件加總公式，依部門/月份彙整銷售數據
'
'作者版權: Dunk
'原始設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestBatchSumIfFormulas()
    Call CreateSumIfFormulaExample
End Sub

' 建立 SUMIF / SUMIFS 公式批次填入示範
Sub CreateSumIfFormulaExample()
    Dim wsData As Worksheet
    Dim wsSummary As Worksheet
    On Error GoTo ErrHandler

    Set wsData = GetOrCreateSumIfSheet(ThisWorkbook, "銷售明細")
    Set wsSummary = GetOrCreateSumIfSheet(ThisWorkbook, "SUMIF彙整")

    Call FillSalesDetailData(wsData)
    Call BuildSumIfSummary(wsSummary, wsData)

    wsSummary.Columns("A:F").AutoFit
    wsSummary.Activate
    MsgBox "SUMIF / SUMIFS 公式已批次填入完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 在彙整表中批次填入 SUMIF / SUMIFS 公式
Private Sub BuildSumIfSummary(ByVal wsSummary As Worksheet, ByVal wsData As Worksheet)
    Dim deptNames(1 To 4) As String
    Dim i As Integer
    Dim dataSheetName As String
    dataSheetName = wsData.Name

    deptNames(1) = "業務部" : deptNames(2) = "行銷部"
    deptNames(3) = "電商部" : deptNames(4) = "通路部"

    ' 標頭
    wsSummary.Range("A1").Value = "部門"
    wsSummary.Range("B1").Value = "總業績(SUMIF)"
    wsSummary.Range("C1").Value = "Q1業績(SUMIFS)"
    wsSummary.Range("D1").Value = "Q2業績(SUMIFS)"
    wsSummary.Range("E1").Value = "線上通路(SUMIFS)"
    wsSummary.Range("F1").Value = "線下通路(SUMIFS)"
    wsSummary.Range("A1:F1").Font.Bold = True

    ' 批次填入各部門公式
    For i = 1 To 4
        wsSummary.Cells(i + 1, 1).Value = deptNames(i)

        ' SUMIF：依部門加總業績
        wsSummary.Cells(i + 1, 2).Formula = "=SUMIF('" & dataSheetName & "'!$B:$B,A" & _
            (i + 1) & ",'" & dataSheetName & "'!$D:$D)"

        ' SUMIFS：部門 + Q1 (1~3月)
        wsSummary.Cells(i + 1, 3).Formula = "=SUMIFS('" & dataSheetName & "'!$D:$D,'" & _
            dataSheetName & "'!$B:$B,A" & (i + 1) & ",'" & dataSheetName & "'!$C:$C,""Q1"")"

        ' SUMIFS：部門 + Q2 (4~6月)
        wsSummary.Cells(i + 1, 4).Formula = "=SUMIFS('" & dataSheetName & "'!$D:$D,'" & _
            dataSheetName & "'!$B:$B,A" & (i + 1) & ",'" & dataSheetName & "'!$C:$C,""Q2"")"

        ' SUMIFS：部門 + 線上通路
        wsSummary.Cells(i + 1, 5).Formula = "=SUMIFS('" & dataSheetName & "'!$D:$D,'" & _
            dataSheetName & "'!$B:$B,A" & (i + 1) & ",'" & dataSheetName & "'!$E:$E,""線上"")"

        ' SUMIFS：部門 + 線下通路
        wsSummary.Cells(i + 1, 6).Formula = "=SUMIFS('" & dataSheetName & "'!$D:$D,'" & _
            dataSheetName & "'!$B:$B,A" & (i + 1) & ",'" & dataSheetName & "'!$E:$E,""線下"")"
    Next i

    ' 格式設定
    wsSummary.Range("B2:F5").NumberFormat = "#,##0"
End Sub

' 填入銷售明細資料
Private Sub FillSalesDetailData(ByVal ws As Worksheet)
    ws.Range("A1:E1").Value = Array("月份", "部門", "季度", "業績", "通路")
    ws.Range("A1:E1").Font.Bold = True

    Dim data As Variant
    data = Array( _
        Array("1月", "業務部", "Q1", 320000, "線下"), _
        Array("1月", "行銷部", "Q1", 180000, "線上"), _
        Array("2月", "業務部", "Q1", 295000, "線下"), _
        Array("2月", "電商部", "Q1", 420000, "線上"), _
        Array("3月", "行銷部", "Q1", 210000, "線上"), _
        Array("3月", "通路部", "Q1", 155000, "線下"), _
        Array("4月", "業務部", "Q2", 380000, "線下"), _
        Array("4月", "電商部", "Q2", 510000, "線上"), _
        Array("5月", "行銷部", "Q2", 260000, "線上"), _
        Array("5月", "通路部", "Q2", 190000, "線下"), _
        Array("6月", "業務部", "Q2", 430000, "線下"), _
        Array("6月", "電商部", "Q2", 580000, "線上") _
    )

    Dim i As Integer
    For i = 0 To UBound(data)
        ws.Cells(i + 2, 1).Value = data(i)(0)
        ws.Cells(i + 2, 2).Value = data(i)(1)
        ws.Cells(i + 2, 3).Value = data(i)(2)
        ws.Cells(i + 2, 4).Value = data(i)(3)
        ws.Cells(i + 2, 5).Value = data(i)(4)
    Next i

    ws.Range("D2:D13").NumberFormat = "#,##0"
    ws.Columns("A:E").AutoFit
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateSumIfSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSumIfSheet = ws
End Function