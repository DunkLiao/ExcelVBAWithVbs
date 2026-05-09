Attribute VB_Name = "BatchFinancialFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchFinancialFormulas
'功能說明: 批次填入財務公式，包含 PMT 月繳、PV 現值、FV 終值、NPV 淨現值
'
'作者版權: Dunk
'原始設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestBatchFinancialFormulas()
    Call CreateFinancialFormulaExample
End Sub

' 建立財務公式批次填入示範
Sub CreateFinancialFormulaExample()
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    Set ws = GetOrCreateFinSheet(ThisWorkbook, "財務公式示範")
    Call BuildLoanPMTTable(ws)
    Call BuildPVFVTable(ws)

    ws.Columns("A:H").AutoFit
    ws.Activate
    MsgBox "財務公式已批次填入完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 建立貸款月繳金額 (PMT) 計算表
Private Sub BuildLoanPMTTable(ByVal ws As Worksheet)
    ' 標題區
    ws.Range("A1").Value = "=== 貸款月繳試算 (PMT) ==="
    ws.Range("A1").Font.Bold = True

    ws.Range("A2").Value = "貸款金額"
    ws.Range("B2").Value = "年利率"
    ws.Range("C2").Value = "年數"
    ws.Range("D2").Value = "月繳金額(PMT)"
    ws.Range("E2").Value = "總繳金額"
    ws.Range("F2").Value = "利息總額"
    ws.Range("A2:F2").Font.Bold = True

    ' 不同貸款條件資料
    Dim loanData As Variant
    loanData = Array( _
        Array(500000, 0.02, 5), _
        Array(1000000, 0.025, 10), _
        Array(2000000, 0.03, 20), _
        Array(3000000, 0.035, 30), _
        Array(500000, 0.015, 3) _
    )

    Dim i As Integer
    For i = 0 To UBound(loanData)
        Dim r As Integer
        r = i + 3
        ws.Cells(r, 1).Value = loanData(i)(0)   ' 貸款金額
        ws.Cells(r, 2).Value = loanData(i)(1)   ' 年利率
        ws.Cells(r, 3).Value = loanData(i)(2)   ' 年數
        ws.Cells(r, 1).NumberFormat = "#,##0"
        ws.Cells(r, 2).NumberFormat = "0.00%"

        ' PMT 公式：=PMT(年利率/12, 年數*12, -貸款金額)
        ws.Cells(r, 4).Formula = "=PMT(B" & r & "/12,C" & r & "*12,-A" & r & ")"
        ws.Cells(r, 4).NumberFormat = "#,##0"

        ' 總繳金額 = 月繳 * 月數
        ws.Cells(r, 5).Formula = "=D" & r & "*C" & r & "*12"
        ws.Cells(r, 5).NumberFormat = "#,##0"

        ' 利息總額 = 總繳 - 貸款金額
        ws.Cells(r, 6).Formula = "=E" & r & "-A" & r
        ws.Cells(r, 6).NumberFormat = "#,##0"
    Next i
End Sub

' 建立 PV (現值) / FV (終值) 計算表
Private Sub BuildPVFVTable(ByVal ws As Worksheet)
    Dim startRow As Integer
    startRow = 10

    ws.Cells(startRow, 1).Value = "=== 投資現值(PV)與終值(FV)試算 ==="
    ws.Cells(startRow, 1).Font.Bold = True

    ws.Cells(startRow + 1, 1).Value = "每期金額"
    ws.Cells(startRow + 1, 2).Value = "年利率"
    ws.Cells(startRow + 1, 3).Value = "期數(年)"
    ws.Cells(startRow + 1, 4).Value = "現值(PV)"
    ws.Cells(startRow + 1, 5).Value = "終值(FV)"
    ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + 1, 5)).Font.Bold = True

    ' 不同投資方案
    Dim pvData As Variant
    pvData = Array( _
        Array(10000, 0.05, 10), _
        Array(20000, 0.04, 15), _
        Array(5000, 0.06, 20), _
        Array(15000, 0.03, 5), _
        Array(8000, 0.055, 12) _
    )

    Dim i As Integer
    For i = 0 To UBound(pvData)
        Dim r As Integer
        r = startRow + 2 + i
        ws.Cells(r, 1).Value = pvData(i)(0)
        ws.Cells(r, 2).Value = pvData(i)(1)
        ws.Cells(r, 3).Value = pvData(i)(2)
        ws.Cells(r, 1).NumberFormat = "#,##0"
        ws.Cells(r, 2).NumberFormat = "0.00%"

        ' PV 公式：每年領回固定金額的現值
        ws.Cells(r, 4).Formula = "=PV(B" & r & ",C" & r & ",-A" & r & ")"
        ws.Cells(r, 4).NumberFormat = "#,##0"

        ' FV 公式：每年存入固定金額的終值
        ws.Cells(r, 5).Formula = "=FV(B" & r & ",C" & r & ",-A" & r & ")"
        ws.Cells(r, 5).NumberFormat = "#,##0"
    Next i
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateFinSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateFinSheet = ws
End Function