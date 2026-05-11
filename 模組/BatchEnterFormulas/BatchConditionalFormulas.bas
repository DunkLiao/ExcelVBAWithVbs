Attribute VB_Name = "BatchConditionalFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchConditionalFormulas
'功能說明: 批次為工作表多個儲存格輸入 IF、巢狀 IF、IFERROR、AND、OR 等條件公式的範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestBatchConditionalFormulas()
    Call InsertBatchConditionalFormulas("批次條件公式")
End Sub

' 批次輸入條件公式
Sub InsertBatchConditionalFormulas(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet

    Set ws = GetOrCreateWorksheetBCF(sheetName)
    ws.Cells.Clear

    ' 填入範例資料
    Call FillConditionalFormulaData(ws)

    ' 設定標題
    ws.Range("A1").Value = "序號"
    ws.Range("B1").Value = "成績"
    ws.Range("C1").Value = "是否及格 (IF)"
    ws.Range("D1").Value = "等級 (巢狀IF)"
    ws.Range("E1").Value = "成績/序號 (IFERROR)"
    ws.Range("F1").Value = "中間段 (AND)"
    ws.Range("G1").Value = "特殊 (OR)"
    ws.Rows(1).Font.Bold = True

    ' 欄 C：IF 公式 — 判斷成績是否及格
    ws.Range("C2:C11").Formula = "=IF(B2>=60,""及格"",""不及格"")"

    ' 欄 D：巢狀 IF — 成績等級分類
    ws.Range("D2:D11").Formula = _
        "=IF(B2>=90,""優"",IF(B2>=80,""良"",IF(B2>=70,""中"",IF(B2>=60,""可"",""差""))))"

    ' 欄 E：IFERROR — 除法保護（避免除以零）
    ws.Range("E2:E11").Formula = "=IFERROR(B2/A2,""N/A"")"

    ' 欄 F：AND 複合條件 — 同時大於 60 且小於 90
    ws.Range("F2:F11").Formula = "=IF(AND(B2>=60,B2<90),""中間段"",""非中間段"")"

    ' 欄 G：OR 複合條件 — 不及格或滿分
    ws.Range("G2:G11").Formula = "=IF(OR(B2<60,B2=100),""特殊注意"",""一般"")"

    ws.UsedRange.Columns.AutoFit

    MsgBox "批次條件公式已輸入完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "輸入條件公式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function GetOrCreateWorksheetBCF(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheetBCF = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetOrCreateWorksheetBCF Is Nothing Then
        Set GetOrCreateWorksheetBCF = ThisWorkbook.Worksheets.Add
        GetOrCreateWorksheetBCF.Name = sheetName
    End If
End Function

Private Sub FillConditionalFormulaData(ByVal ws As Worksheet)
    Dim scores(1 To 10) As Integer
    scores(1) = 95 : scores(2) = 82 : scores(3) = 74
    scores(4) = 63 : scores(5) = 55 : scores(6) = 88
    scores(7) = 100 : scores(8) = 45 : scores(9) = 71
    scores(10) = 60

    Dim r As Integer
    For r = 1 To 10
        ws.Cells(r + 1, 1).Value = r
        ws.Cells(r + 1, 2).Value = scores(r)
    Next r
End Sub
