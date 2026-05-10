Option Explicit
'*************************************************************************************
'模組名稱: BudgetVarianceFormatting
'功能說明: 建立預算差異區間的條件格式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Public Sub ApplyBudgetVarianceFormatting()
    Dim ws As Worksheet
    Dim varianceRange As Range
    Dim fc As FormatCondition

    On Error GoTo ErrHandler

    Set ws = GetOrCreateBudgetSheet("預算差異格式範例")
    ws.Cells.Clear
    Call FillBudgetData(ws)

    Set varianceRange = ws.Range("D2:D9")
    varianceRange.FormatConditions.Delete

    Set fc = varianceRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="=-0.1")
    With fc
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
    End With

    Set fc = varianceRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="=-0.1", Formula2:="=0.1")
    With fc
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With

    Set fc = varianceRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0.1")
    With fc
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 101, 0)
    End With

    ws.Range("D2:D9").NumberFormat = "0.0%"
    ws.Columns("A:D").AutoFit
    MsgBox "預算差異區間條件格式已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立預算差異條件格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillBudgetData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("科目", "預算", "實際", "差異率")
    ws.Range("A2:C2").Value = Array("人事費", 420000, 438000)
    ws.Range("A3:C3").Value = Array("租金", 180000, 180000)
    ws.Range("A4:C4").Value = Array("廣告費", 90000, 114000)
    ws.Range("A5:C5").Value = Array("差旅費", 65000, 51000)
    ws.Range("A6:C6").Value = Array("設備費", 120000, 132500)
    ws.Range("A7:C7").Value = Array("教育訓練", 45000, 39000)
    ws.Range("A8:C8").Value = Array("顧問費", 80000, 96000)
    ws.Range("A9:C9").Value = Array("雜項", 30000, 28000)
    ws.Range("D2:D9").FormulaR1C1 = "=(RC[-1]-RC[-2])/RC[-2]"
End Sub

Private Function GetOrCreateBudgetSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateBudgetSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateBudgetSheet Is Nothing Then
        Set GetOrCreateBudgetSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateBudgetSheet.Name = sheetName
    End If
End Function
