Attribute VB_Name = "BatchVlookupFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchVlookupFormulas
'功能說明: 批次在多個儲存格中輸入 VLOOKUP 公式，從對照表自動查詢對應值
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestBatchVlookupFormulas()
    Call CreateVlookupBatchExample
End Sub

' 建立 VLOOKUP 批次公式範例
Sub CreateVlookupBatchExample()
    On Error GoTo ErrorHandler

    Dim wb As Workbook
    Dim lookupWs As Worksheet
    Dim dataWs As Worksheet

    Set wb = ThisWorkbook
    Set lookupWs = GetOrCreateSheet(wb, "對照表")
    Set dataWs = GetOrCreateSheet(wb, "VLOOKUP批次範例")

    Call FillLookupTable(lookupWs)
    Call FillMainData(dataWs)
    Call BatchEnterVlookupFormulas(dataWs, lookupWs)

    dataWs.Activate
    MsgBox "VLOOKUP 批次公式已輸入完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "批次輸入 VLOOKUP 公式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 批次輸入 VLOOKUP 公式
Private Sub BatchEnterVlookupFormulas(ByVal dataWs As Worksheet, ByVal lookupWs As Worksheet)
    Dim lastRow As Long
    Dim i As Long
    Dim lookupRef As String

    lastRow = dataWs.Cells(dataWs.Rows.Count, 1).End(xlUp).Row

    lookupRef = "'" & lookupWs.Name & "'!$A:$C"

    For i = 2 To lastRow
        ' 欄 C：從對照表查詢部門名稱（第2欄）
        dataWs.Cells(i, 3).Formula = _
            "=IFERROR(VLOOKUP(" & _
            dataWs.Cells(i, 1).Address(False, False) & _
            "," & lookupRef & ",2,FALSE),""查無資料"")"

        ' 欄 D：從對照表查詢職稱（第3欄）
        dataWs.Cells(i, 4).Formula = _
            "=IFERROR(VLOOKUP(" & _
            dataWs.Cells(i, 1).Address(False, False) & _
            "," & lookupRef & ",3,FALSE),""查無資料"")"
    Next i

    dataWs.Columns("A:D").AutoFit
End Sub

' 填入對照表資料（員工代碼 → 部門、職稱）
Private Sub FillLookupTable(ByVal ws As Worksheet)
    ws.Range("A1").Value = "員工代碼"
    ws.Range("B1").Value = "部門"
    ws.Range("C1").Value = "職稱"
    ws.Range("A1:C1").Font.Bold = True

    Dim data As Variant
    data = Array( _
        Array("E001", "業務部", "業務專員"), _
        Array("E002", "研發部", "工程師"), _
        Array("E003", "行政部", "行政助理"), _
        Array("E004", "業務部", "業務經理"), _
        Array("E005", "研發部", "資深工程師"), _
        Array("E006", "財務部", "會計師"), _
        Array("E007", "人資部", "人資專員"), _
        Array("E008", "行政部", "行政主任") _
    )

    Dim i As Integer
    For i = 0 To 7
        ws.Cells(i + 2, 1).Value = data(i)(0)
        ws.Cells(i + 2, 2).Value = data(i)(1)
        ws.Cells(i + 2, 3).Value = data(i)(2)
    Next i

    ws.Columns("A:C").AutoFit
End Sub

' 填入主資料（員工代碼和姓名，待查詢部門和職稱）
Private Sub FillMainData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "員工代碼"
    ws.Range("B1").Value = "姓名"
    ws.Range("C1").Value = "部門"
    ws.Range("D1").Value = "職稱"
    ws.Range("A1:D1").Font.Bold = True

    ws.Range("C1").Interior.Color = RGB(255, 242, 204)
    ws.Range("D1").Interior.Color = RGB(255, 242, 204)

    Dim employees As Variant
    employees = Array( _
        Array("E003", "陳小華"), _
        Array("E007", "林美君"), _
        Array("E001", "張志明"), _
        Array("E005", "王建豪"), _
        Array("E009", "劉雅芳"), _
        Array("E002", "李俊傑"), _
        Array("E006", "黃淑惠"), _
        Array("E004", "吳志遠") _
    )

    Dim i As Integer
    For i = 0 To 7
        ws.Cells(i + 2, 1).Value = employees(i)(0)
        ws.Cells(i + 2, 2).Value = employees(i)(1)
    Next i
End Sub

' 取得或建立工作表，並清除內容
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
