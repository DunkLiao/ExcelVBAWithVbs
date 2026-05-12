Attribute VB_Name = "BatchXlookupFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchXlookupFormulas
'功能說明: 示範如何批次在指定欄位插入 XLOOKUP 公式，適用於 Excel 2019+ 及 Microsoft 365
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestBatchXlookupFormulas()
    Call CreateXlookupExample
End Sub

' 建立 XLOOKUP 公式範例
Sub CreateXlookupExample()
    On Error GoTo ErrorHandler

    Dim wsLookup As Worksheet
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim codes As Variant
    Dim lookupRange As String
    Dim notFoundMsg As String
    Dim lookupVal As String

    Set wsData = GetOrCreateSheet(ThisWorkbook, "產品對照表")
    Call FillProductTable(wsData)

    Set wsLookup = GetOrCreateSheet(ThisWorkbook, "XLOOKUP批次查詢")

    wsLookup.Range("A1").Value = "產品代碼"
    wsLookup.Range("B1").Value = "產品名稱（XLOOKUP）"
    wsLookup.Range("C1").Value = "單價（XLOOKUP）"
    wsLookup.Range("D1").Value = "類別（XLOOKUP）"
    wsLookup.Range("A1:D1").Font.Bold = True

    codes = Array("P001", "P003", "P005", "P002", "P004", "P999")
    For i = 0 To UBound(codes)
        wsLookup.Cells(i + 2, 1).Value = codes(i)
    Next i

    lastRow = UBound(codes) + 2
    notFoundMsg = """查無資料"""
    lookupRange = "產品對照表!$A$2:$A$6"

    For i = 2 To lastRow
        lookupVal = wsLookup.Cells(i, 1).Address(False, False)

        wsLookup.Cells(i, 2).Formula = "=XLOOKUP(" & lookupVal & "," & _
            lookupRange & ",產品對照表!$B$2:$B$6," & notFoundMsg & ")"

        wsLookup.Cells(i, 3).Formula = "=XLOOKUP(" & lookupVal & "," & _
            lookupRange & ",產品對照表!$C$2:$C$6," & notFoundMsg & ")"

        wsLookup.Cells(i, 4).Formula = "=XLOOKUP(" & lookupVal & "," & _
            lookupRange & ",產品對照表!$D$2:$D$6," & notFoundMsg & ")"
    Next i

    wsLookup.Columns("A:D").AutoFit

    MsgBox "XLOOKUP 批次公式已插入完成！" & vbCrLf & _
           "注意：XLOOKUP 需要 Excel 2019+ 或 Microsoft 365。", _
           vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立 XLOOKUP 公式時發生錯誤：" & Err.Description & vbCrLf & _
           "請確認 Excel 版本支援 XLOOKUP 函數。", vbExclamation, "錯誤"
End Sub

' 填入產品對照表
Private Sub FillProductTable(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品代碼"
    ws.Range("B1").Value = "產品名稱"
    ws.Range("C1").Value = "單價"
    ws.Range("D1").Value = "類別"
    ws.Range("A1:D1").Font.Bold = True

    ws.Range("A2").Value = "P001"
    ws.Range("B2").Value = "筆記型電腦"
    ws.Range("C2").Value = 32000
    ws.Range("D2").Value = "電子"

    ws.Range("A3").Value = "P002"
    ws.Range("B3").Value = "無線滑鼠"
    ws.Range("C3").Value = 850
    ws.Range("D3").Value = "電子"

    ws.Range("A4").Value = "P003"
    ws.Range("B4").Value = "辦公椅"
    ws.Range("C4").Value = 4500
    ws.Range("D4").Value = "家具"

    ws.Range("A5").Value = "P004"
    ws.Range("B5").Value = "A4 影印紙"
    ws.Range("C5").Value = 120
    ws.Range("D5").Value = "文具"

    ws.Range("A6").Value = "P005"
    ws.Range("B6").Value = "鍵盤"
    ws.Range("C6").Value = 1200
    ws.Range("D6").Value = "電子"

    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
