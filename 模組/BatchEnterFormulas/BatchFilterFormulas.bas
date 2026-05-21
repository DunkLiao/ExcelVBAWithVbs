Option Explicit
Attribute VB_Name = "BatchFilterFormulas"
'*************************************************************************************
'模組名稱: BatchFilterFormulas
'功能說明: 批次插入 FILTER 動態陣列函數公式，依不同條件篩選資料至各目標位置
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestBatchFilterFormulas()
    Call CreateFilterFormulaSheet("FILTER公式範例")
End Sub

Sub CreateFilterFormulaSheet(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateBFFSheet(sheetName)
    ws.Cells.Clear

    Call FillFilterSourceData(ws)
    Call InsertFilterFormulas(ws)

    ws.Columns.AutoFit
    MsgBox "FILTER 函數公式已批次插入完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "插入 FILTER 公式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillFilterSourceData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("姓名", "部門", "職等", "薪資")
    ws.Range("A2:D2").Value = Array("張大明", "業務部", 3, 45000)
    ws.Range("A3:D3").Value = Array("李小華", "人事部", 2, 38000)
    ws.Range("A4:D4").Value = Array("王美麗", "業務部", 4, 55000)
    ws.Range("A5:D5").Value = Array("陳志偉", "資訊部", 3, 48000)
    ws.Range("A6:D6").Value = Array("林怡君", "人事部", 1, 32000)
    ws.Range("A7:D7").Value = Array("吳建國", "資訊部", 5, 72000)
    ws.Range("A8:D8").Value = Array("黃淑芬", "業務部", 2, 40000)

    ws.Range("F1").Value = "=== 業務部 FILTER ==="
    ws.Range("H1").Value = "=== 薪資 >= 45000 FILTER ==="
    ws.Range("J1").Value = "=== 資訊部 FILTER ==="
End Sub

Private Sub InsertFilterFormulas(ByVal ws As Worksheet)
    ' 篩選業務部員工
    ws.Range("F2").Formula = _
        "=FILTER(A2:D8,B2:B8=" & Chr(34) & "業務部" & Chr(34) & _
        "," & Chr(34) & "查無資料" & Chr(34) & ")"

    ' 篩選薪資大於等於 45000 的員工
    ws.Range("H2").Formula = _
        "=FILTER(A2:D8,D2:D8>=45000," & Chr(34) & "查無資料" & Chr(34) & ")"

    ' 篩選資訊部員工
    ws.Range("J2").Formula = _
        "=FILTER(A2:D8,B2:B8=" & Chr(34) & "資訊部" & Chr(34) & _
        "," & Chr(34) & "查無資料" & Chr(34) & ")"
End Sub

Private Function GetOrCreateBFFSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateBFFSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateBFFSheet Is Nothing Then
        Set GetOrCreateBFFSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateBFFSheet.Name = sheetName
    End If
End Function
