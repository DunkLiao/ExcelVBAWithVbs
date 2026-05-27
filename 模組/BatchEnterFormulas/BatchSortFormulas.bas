Option Explicit
Attribute VB_Name = "BatchSortFormulas"
'*************************************************************************************
'模組名稱: 批次輸入 SORT 公式
'功能說明: 以 VBA 批次寫入 Excel 365 動態陣列 SORT 函數，
'          示範依單欄升冪、降冪等不同方式排序
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub TestBatchSortFormulas()
    Call CreateBatchSortFormulas("SORT公式範例")
End Sub

Sub CreateBatchSortFormulas(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateWsSort(sheetName)
    ws.Cells.Clear

    Call FillSortSourceData(ws)
    Call WriteSortFormulas(ws)

    ws.Columns("A:K").AutoFit
    MsgBox "SORT 公式範例已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立 SORT 公式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function GetOrCreateWsSort(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWsSort = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWsSort Is Nothing Then
        Set GetOrCreateWsSort = ThisWorkbook.Worksheets.Add
        GetOrCreateWsSort.Name = sheetName
    End If
End Function

Private Sub FillSortSourceData(ByVal ws As Worksheet)
    ' 原始資料 A1:C7
    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "業績"

    ws.Range("A2").Value = "王大明"
    ws.Range("B2").Value = "業務部"
    ws.Range("C2").Value = 320000

    ws.Range("A3").Value = "李小華"
    ws.Range("B3").Value = "財務部"
    ws.Range("C3").Value = 150000

    ws.Range("A4").Value = "張美玲"
    ws.Range("B4").Value = "業務部"
    ws.Range("C4").Value = 480000

    ws.Range("A5").Value = "陳建宏"
    ws.Range("B5").Value = "研發部"
    ws.Range("C5").Value = 210000

    ws.Range("A6").Value = "林淑芬"
    ws.Range("B6").Value = "業務部"
    ws.Range("C6").Value = 390000

    ws.Range("A7").Value = "黃志偉"
    ws.Range("B7").Value = "財務部"
    ws.Range("C7").Value = 175000
End Sub

Private Sub WriteSortFormulas(ByVal ws As Worksheet)
    ' 依業績降冪排序（第3欄，-1=降冪）
    ws.Range("E1").Value = "依業績降冪"
    ws.Range("E2").Formula = "=SORT(A2:C7,3,-1)"

    ' 依姓名升冪排序（第1欄，1=升冪）
    ws.Range("I1").Value = "依姓名升冪"
    ws.Range("I2").Formula = "=SORT(A2:C7,1,1)"

    ' 說明文字
    ws.Range("E9").Value = "說明：SORT(陣列, 排序欄號, 1升冪/-1降冪)"
End Sub
