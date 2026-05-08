Option Explicit

' 比對兩份清單的鍵值，標示只存在於左表或右表的資料。
Public Sub ReconcileTwoListsExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet

    Set ws = GetOrCreateReconcileWorksheet("雙清單比對範例")
    ws.Cells.Clear
    Call FillReconcileData(ws)
    Call ReconcileTwoLists(ws)

    ws.Columns("A:G").AutoFit
    MsgBox "雙清單比對完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "比對清單失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub ReconcileTwoLists(ByVal ws As Worksheet)
    Dim leftLastRow As Long
    Dim rightLastRow As Long
    Dim resultRow As Long
    Dim rowIndex As Long
    Dim foundCell As Range

    leftLastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    rightLastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    ws.Range("G1").Value = "差異說明"
    resultRow = 2

    For rowIndex = 2 To leftLastRow
        Set foundCell = ws.Range("D2:D" & rightLastRow).Find(What:=ws.Cells(rowIndex, "A").Value, LookIn:=xlValues, LookAt:=xlWhole)
        If foundCell Is Nothing Then
            ws.Cells(resultRow, "G").Value = "左表有，右表沒有：" & ws.Cells(rowIndex, "A").Value
            resultRow = resultRow + 1
        End If
    Next rowIndex

    For rowIndex = 2 To rightLastRow
        Set foundCell = ws.Range("A2:A" & leftLastRow).Find(What:=ws.Cells(rowIndex, "D").Value, LookIn:=xlValues, LookAt:=xlWhole)
        If foundCell Is Nothing Then
            ws.Cells(resultRow, "G").Value = "右表有，左表沒有：" & ws.Cells(rowIndex, "D").Value
            resultRow = resultRow + 1
        End If
    Next rowIndex
End Sub

Private Sub FillReconcileData(ByVal ws As Worksheet)
    ws.Range("A1:B1").Value = Array("左表編號", "左表金額")
    ws.Range("A2:B2").Value = Array("A001", 1200)
    ws.Range("A3:B3").Value = Array("A002", 1800)
    ws.Range("A4:B4").Value = Array("A003", 1500)
    ws.Range("A5:B5").Value = Array("A004", 2100)
    ws.Range("D1:E1").Value = Array("右表編號", "右表金額")
    ws.Range("D2:E2").Value = Array("A001", 1200)
    ws.Range("D3:E3").Value = Array("A003", 1500)
    ws.Range("D4:E4").Value = Array("A004", 2100)
    ws.Range("D5:E5").Value = Array("A005", 990)
End Sub

Private Function GetOrCreateReconcileWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateReconcileWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateReconcileWorksheet Is Nothing Then
        Set GetOrCreateReconcileWorksheet = ThisWorkbook.Worksheets.Add
        GetOrCreateReconcileWorksheet.Name = sheetName
    End If
End Function