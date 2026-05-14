Attribute VB_Name = "BatchUniqueFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchUniqueFormulas
'功能說明: 批次在工作表中寫入 UNIQUE 函數及相關動態陣列公式的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點
Sub TestBatchUniqueFormulas()
    Call CreateUniqueFormulaSheet("UNIQUE函數範例")
End Sub

' 建立 UNIQUE 函數範例工作表
' sheetName: 目標工作表名稱
Sub CreateUniqueFormulaSheet(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateUniqueSheet(sheetName)
    ws.Cells.Clear

    Call FillUniqueSourceData(ws)

    ws.Range("D1").Value = "UNIQUE 系列公式示範"
    ws.Range("D1").Font.Bold = True
    ws.Range("D1").Font.Size = 13

    ws.Range("D3").Value = "=UNIQUE(A2:A11)"
    ws.Range("E3").Value = "傳回不重複的部門清單"

    ws.Range("D10").Value = "=UNIQUE(A2:A11,,TRUE)"
    ws.Range("E10").Value = "只傳回恰好出現一次的部門"

    ws.Range("D17").Value = "=SORT(UNIQUE(A2:A11))"
    ws.Range("E17").Value = "排序後的不重複部門清單"

    ws.Range("D24").Value = "=COUNTA(UNIQUE(A2:A11))"
    ws.Range("E24").Value = "不重複部門的數量"

    ws.Range("D26").Value = "=UNIQUE(A2:B11)"
    ws.Range("E26").Value = "部門+職稱的不重複組合"

    ws.Columns("A:F").AutoFit

    MsgBox "UNIQUE 函數範例已建立完成！" & Chr(10) & _
           "注意：UNIQUE 函數需要 Excel 365 或 Excel 2021 以上版本。", _
           vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立 UNIQUE 公式範例時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 填入來源資料
Private Sub FillUniqueSourceData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("部門", "職稱", "薪資")
    ws.Range("A1:C1").Font.Bold = True

    ws.Range("A2:C2").Value = Array("業務部", "專員", 35000)
    ws.Range("A3:C3").Value = Array("行銀部", "主任", 55000)
    ws.Range("A4:C4").Value = Array("技術部", "工程師", 65000)
    ws.Range("A5:C5").Value = Array("業務部", "組長", 48000)
    ws.Range("A6:C6").Value = Array("行政部", "助理", 30000)
    ws.Range("A7:C7").Value = Array("技術部", "高級工程師", 85000)
    ws.Range("A8:C8").Value = Array("行銀部", "設計師", 52000)
    ws.Range("A9:C9").Value = Array("業務部", "專員", 36000)
    ws.Range("A10:C10").Value = Array("行政部", "主任", 58000)
    ws.Range("A11:C11").Value = Array("技術部", "工程師", 67000)

    ws.Range("C2:C11").NumberFormat = "#,##0"
End Sub

' 取得或建立工作表
Private Function GetOrCreateUniqueSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateUniqueSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateUniqueSheet Is Nothing Then
        Set GetOrCreateUniqueSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateUniqueSheet.Name = sheetName
    End If
End Function
