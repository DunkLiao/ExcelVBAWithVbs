Attribute VB_Name = "FinanceFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: FinanceFormulaExample
'功能描述: 在 Excel 中示範財務公式的使用範例
'          包含 PMT、PV、FV、RATE、NPER、IRR、NPV 等公式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

Sub TestFinanceFormula()
    Call CreateFinanceFormulaExample("財務公式範例")
End Sub

Sub CreateFinanceFormulaExample(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = "===== 貸款計算 (PMT) ====="
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "貸款金額 (元)"
    ws.Range("B2").Value = 1000000
    ws.Range("A3").Value = "年利率"
    ws.Range("B3").Value = 0.03
    ws.Range("B3").NumberFormat = "0.00%"
    ws.Range("A4").Value = "還款期數 (月)"
    ws.Range("B4").Value = 240
    ws.Range("A5").Value = "每月還款金額"
    ws.Range("B5").Formula = "=PMT(B3/12,B4,-B2)"
    ws.Range("B5").NumberFormat = "#,##0.00"
    ws.Range("B5").Interior.Color = RGB(198, 239, 206)
    ws.Range("A7").Value = "===== 現值計算 (PV) ====="
    ws.Range("A7").Font.Bold = True
    ws.Range("A8").Value = "年利率"
    ws.Range("B8").Value = 0.05
    ws.Range("B8").NumberFormat = "0.00%"
    ws.Range("A9").Value = "期數"
    ws.Range("B9").Value = 10
    ws.Range("A10").Value = "每期現金流"
    ws.Range("B10").Value = 50000
    ws.Range("A11").Value = "現值"
    ws.Range("B11").Formula = "=PV(B8,B9,-B10)"
    ws.Range("B11").NumberFormat = "#,##0.00"
    ws.Range("B11").Interior.Color = RGB(198, 239, 206)
    ws.Range("A13").Value = "===== 終值計算 (FV) ====="
    ws.Range("A13").Font.Bold = True
    ws.Range("A14").Value = "年利率"
    ws.Range("B14").Value = 0.04
    ws.Range("B14").NumberFormat = "0.00%"
    ws.Range("A15").Value = "期數"
    ws.Range("B15").Value = 12
    ws.Range("A16").Value = "每期存入"
    ws.Range("B16").Value = 10000
    ws.Range("A17").Value = "終值"
    ws.Range("B17").Formula = "=FV(B14/12,B15,-B16)"
    ws.Range("B17").NumberFormat = "#,##0.00"
    ws.Range("B17").Interior.Color = RGB(198, 239, 206)
    ws.Range("A19").Value = "===== NPV 與 IRR ====="
    ws.Range("A19").Font.Bold = True
    ws.Range("A20").Value = "折現率"
    ws.Range("B20").Value = 0.08
    ws.Range("B20").NumberFormat = "0.00%"
    ws.Range("A21").Value = "初始投資"
    ws.Range("B21").Value = -100000
    ws.Range("A22").Value = "第1年現金流"
    ws.Range("B22").Value = 30000
    ws.Range("A23").Value = "第2年現金流"
    ws.Range("B23").Value = 40000
    ws.Range("A24").Value = "第3年現金流"
    ws.Range("B24").Value = 50000
    ws.Range("A25").Value = "第4年現金流"
    ws.Range("B25").Value = 30000
    ws.Range("A26").Value = "NPV 淨現值"
    ws.Range("B26").Formula = "=B21+NPV(B20,B22:B25)"
    ws.Range("B26").NumberFormat = "#,##0.00"
    ws.Range("B26").Interior.Color = RGB(255, 235, 156)
    ws.Range("A27").Value = "IRR 內部報酬率"
    ws.Range("B27").Formula = "=IRR(B21:B25)"
    ws.Range("B27").NumberFormat = "0.00%"
    ws.Range("B27").Interior.Color = RGB(255, 235, 156)
    ws.Columns("A:B").AutoFit
    MsgBox "財務公式範例已建立完成！", vbInformation, "完成"
End Sub
