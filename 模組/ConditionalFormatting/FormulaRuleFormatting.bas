Option Explicit
'*************************************************************************************
'模組名稱: FormulaRuleFormatting
'功能說明: 使用自訂公式建立條件格式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Public Sub ApplyFormulaRuleFormatting()
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim fc As FormatCondition

    On Error GoTo ErrHandler

    Set ws = GetOrCreateFormulaSheet("公式規則格式範例")
    ws.Cells.Clear
    Call FillFormulaRuleData(ws)

    Set targetRange = ws.Range("A2:E10")
    targetRange.FormatConditions.Delete

    Set fc = targetRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($D2<>"""",$D2<$E2)")
    With fc
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With

    Set fc = targetRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($D2<>"""",$D2>=$E2)")
    With fc
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With

    ws.Columns("A:E").AutoFit
    MsgBox "自訂公式條件格式已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立公式規則條件格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillFormulaRuleData(ByVal ws As Worksheet)
    ws.Range("A1:E1").Value = Array("案件", "負責人", "狀態", "完成日", "期限")
    ws.Range("A2:E2").Value = Array("採購申請", "王小明", "完成", Date - 4, Date - 5)
    ws.Range("A3:E3").Value = Array("合約審查", "李小華", "完成", Date - 1, Date - 3)
    ws.Range("A4:E4").Value = Array("設備驗收", "陳美玲", "進行中", "", Date + 2)
    ws.Range("A5:E5").Value = Array("費用核銷", "張志強", "完成", Date + 1, Date)
    ws.Range("A6:E6").Value = Array("文件歸檔", "林雅婷", "完成", Date - 8, Date - 8)
    ws.Range("A7:E7").Value = Array("教育訓練", "周建宏", "進行中", "", Date + 7)
    ws.Range("A8:E8").Value = Array("系統盤點", "吳佩君", "完成", Date - 2, Date - 1)
    ws.Range("A9:E9").Value = Array("客訴回覆", "許家豪", "完成", Date, Date - 1)
    ws.Range("A10:E10").Value = Array("報表發佈", "黃怡君", "進行中", "", Date + 3)
    ws.Range("D2:E10").NumberFormat = "yyyy/m/d"
End Sub

Private Function GetOrCreateFormulaSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateFormulaSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateFormulaSheet Is Nothing Then
        Set GetOrCreateFormulaSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateFormulaSheet.Name = sheetName
    End If
End Function
