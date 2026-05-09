Attribute VB_Name = "BatchPercentFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchPercentFormulas
'功能說明: 批次填入百分比計算公式，包含佔比、成長率、達成率、毛利率等
'
'作者版權: Dunk
'原始設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestBatchPercentFormulas()
    Call CreatePercentFormulaExample
End Sub

' 建立百分比公式批次填入示範
Sub CreatePercentFormulaExample()
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    Set ws = GetOrCreatePctSheet(ThisWorkbook, "百分比公式示範")
    Call FillSalesTargetData(ws)
    Call BatchEnterPercentFormulas(ws)

    ws.Columns("A:I").AutoFit
    ws.Activate
    MsgBox "百分比計算公式已批次填入完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次填入各種百分比計算公式
Private Sub BatchEnterPercentFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 欄位標頭
    ws.Range("F1").Value = "佔總業績%"
    ws.Range("G1").Value = "目標達成率%"
    ws.Range("H1").Value = "與去年成長率%"
    ws.Range("I1").Value = "毛利率%"
    ws.Range("F1:I1").Font.Bold = True

    ' 計算總業績（用於佔比）
    Dim totalFormula As String
    totalFormula = "SUM($B$2:$B$" & lastRow & ")"

    For i = 2 To lastRow
        ' 佔總業績比例
        ws.Cells(i, 6).Formula = "=IFERROR(B" & i & "/" & totalFormula & ",0)"
        ws.Cells(i, 6).NumberFormat = "0.00%"

        ' 目標達成率 = 實際 / 目標
        ws.Cells(i, 7).Formula = "=IFERROR(B" & i & "/C" & i & ",0)"
        ws.Cells(i, 7).NumberFormat = "0.00%"

        ' 與去年成長率 = (今年 - 去年) / 去年
        ws.Cells(i, 8).Formula = "=IFERROR((B" & i & "-D" & i & ")/D" & i & ",0)"
        ws.Cells(i, 8).NumberFormat = "0.00%"

        ' 毛利率 = (業績 - 成本) / 業績
        ws.Cells(i, 9).Formula = "=IFERROR((B" & i & "-E" & i & ")/B" & i & ",0)"
        ws.Cells(i, 9).NumberFormat = "0.00%"
    Next i

    ' 合計列
    Dim sumRow As Long
    sumRow = lastRow + 1
    ws.Cells(sumRow, 1).Value = "合計"
    ws.Cells(sumRow, 1).Font.Bold = True

    ws.Cells(sumRow, 2).Formula = "=SUM(B2:B" & lastRow & ")"
    ws.Cells(sumRow, 2).NumberFormat = "#,##0"
    ws.Cells(sumRow, 3).Formula = "=SUM(C2:C" & lastRow & ")"
    ws.Cells(sumRow, 3).NumberFormat = "#,##0"
    ws.Cells(sumRow, 4).Formula = "=SUM(D2:D" & lastRow & ")"
    ws.Cells(sumRow, 4).NumberFormat = "#,##0"
    ws.Cells(sumRow, 5).Formula = "=SUM(E2:E" & lastRow & ")"
    ws.Cells(sumRow, 5).NumberFormat = "#,##0"

    ' 合計列達成率與毛利率
    ws.Cells(sumRow, 7).Formula = "=IFERROR(B" & sumRow & "/C" & sumRow & ",0)"
    ws.Cells(sumRow, 7).NumberFormat = "0.00%"
    ws.Cells(sumRow, 9).Formula = "=IFERROR((B" & sumRow & "-E" & sumRow & ")/B" & sumRow & ",0)"
    ws.Cells(sumRow, 9).NumberFormat = "0.00%"
    ws.Range(ws.Cells(sumRow, 1), ws.Cells(sumRow, 9)).Font.Bold = True
End Sub

' 填入各部門業績、目標、去年業績、成本資料
Private Sub FillSalesTargetData(ByVal ws As Worksheet)
    ws.Range("A1:E1").Value = Array("部門", "今年業績", "目標金額", "去年業績", "成本")
    ws.Range("A1:E1").Font.Bold = True

    Dim data As Variant
    data = Array( _
        Array("業務一組", 3850000, 3500000, 3200000, 2310000), _
        Array("業務二組", 2980000, 3200000, 2750000, 1788000), _
        Array("行銷部門", 1560000, 1500000, 1320000, 936000), _
        Array("電商部門", 4200000, 4000000, 3600000, 2520000), _
        Array("通路部門", 2750000, 2800000, 2500000, 1650000), _
        Array("直銷團隊", 1980000, 2000000, 1800000, 1188000) _
    )

    Dim i As Integer
    For i = 0 To UBound(data)
        Dim r As Integer
        r = i + 2
        ws.Cells(r, 1).Value = data(i)(0)
        ws.Cells(r, 2).Value = data(i)(1)
        ws.Cells(r, 3).Value = data(i)(2)
        ws.Cells(r, 4).Value = data(i)(3)
        ws.Cells(r, 5).Value = data(i)(4)
    Next i

    ws.Range("B2:E7").NumberFormat = "#,##0"
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreatePctSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreatePctSheet = ws
End Function