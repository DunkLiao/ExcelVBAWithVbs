Attribute VB_Name = "DynamicArrayFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: DynamicArrayFormulaExample
'功能說明: 在 Excel 中建立動態陣列公式（UNIQUE、SORT、FILTER）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestDynamicArrayFormula()
    Call CreateDynamicArrayExample
End Sub

' 建立動態陣列公式範例（需要 Excel 365 或 Excel 2021+）
Sub CreateDynamicArrayExample()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "動態陣列公式範例")

    Call FillSalesData(ws)
    Call EnterDynamicFormulas(ws)

    ws.Activate
    MsgBox "動態陣列公式已建立完成！" & vbCrLf & _
           "（需要 Excel 365 或 Excel 2021 以上版本才支援）", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立動態陣列公式時發生錯誤：" & Err.Description & vbCrLf & _
           "請確認您的 Excel 版本是否支援動態陣列函數。", vbExclamation, "錯誤"
End Sub

' 輸入動態陣列公式
Private Sub EnterDynamicFormulas(ByVal ws As Worksheet)
    ' UNIQUE 函數：擷取不重複的業務員名單
    ws.Range("F1").Value = "不重複業務員（UNIQUE）"
    ws.Range("F1").Font.Bold = True
    ws.Range("F2").Formula2 = "=UNIQUE(B2:B16)"

    ' SORT 函數：依金額排序
    ws.Range("H1").Value = "依金額降冪排序（SORT）"
    ws.Range("H1").Font.Bold = True
    ws.Range("H2").Formula2 = "=SORT(A2:C16,3,-1)"

    ' FILTER 函數：篩選金額大於 10000 的紀錄
    ws.Range("L1").Value = "金額 > 10000（FILTER）"
    ws.Range("L1").Font.Bold = True
    ws.Range("L2").Formula2 = "=FILTER(A2:C16,C2:C16>10000,""無符合資料"")"

    ws.Columns("F:N").AutoFit
End Sub

' 填入銷售範例資料
Private Sub FillSalesData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "業務員"
    ws.Range("C1").Value = "銷售金額"
    ws.Range("A1:C1").Font.Bold = True

    Dim data As Variant
    data = Array( _
        Array("1月", "張小明", 12000), _
        Array("1月", "李美麗", 8500), _
        Array("2月", "張小明", 15000), _
        Array("2月", "王大同", 9800), _
        Array("2月", "李美麗", 11200), _
        Array("3月", "張小明", 7600), _
        Array("3月", "王大同", 13500), _
        Array("3月", "陳建國", 10200), _
        Array("4月", "李美麗", 16800), _
        Array("4月", "陳建國", 5400), _
        Array("4月", "王大同", 12100), _
        Array("5月", "張小明", 19500), _
        Array("5月", "陳建國", 8700), _
        Array("6月", "李美麗", 14300), _
        Array("6月", "王大同", 11700) _
    )

    Dim i As Integer
    For i = 0 To 14
        ws.Cells(i + 2, 1).Value = data(i)(0)
        ws.Cells(i + 2, 2).Value = data(i)(1)
        ws.Cells(i + 2, 3).Value = data(i)(2)
    Next i

    ws.Columns("A:C").AutoFit
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
