Attribute VB_Name = "BatchIndexMatchFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchIndexMatchFormulas
'功能說明: 批次填入 INDEX+MATCH 雙向查找公式，取代 VLOOKUP 實現更靈活的資料查詢
'
'作者版權: Dunk
'原始設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestBatchIndexMatchFormulas()
    Call CreateIndexMatchFormulaExample
End Sub

' 建立 INDEX/MATCH 公式批次填入示範
Sub CreateIndexMatchFormulaExample()
    Dim wsRef As Worksheet
    Dim wsQuery As Worksheet
    On Error GoTo ErrHandler

    Set wsRef = GetOrCreateIMSheet(ThisWorkbook, "商品主檔")
    Set wsQuery = GetOrCreateIMSheet(ThisWorkbook, "INDEX查詢")

    Call FillProductMasterData(wsRef)
    Call BuildIndexMatchQuery(wsQuery, wsRef)

    wsQuery.Columns("A:F").AutoFit
    wsQuery.Activate
    MsgBox "INDEX + MATCH 公式已批次填入完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次填入 INDEX+MATCH 查詢公式
Private Sub BuildIndexMatchQuery(ByVal wsQuery As Worksheet, ByVal wsRef As Worksheet)
    Dim refName As String
    Dim i As Integer
    refName = wsRef.Name

    ' 查詢標頭
    wsQuery.Range("A1").Value = "查詢商品ID"
    wsQuery.Range("B1").Value = "商品名稱"
    wsQuery.Range("C1").Value = "類別"
    wsQuery.Range("D1").Value = "單價"
    wsQuery.Range("E1").Value = "庫存量"
    wsQuery.Range("F1").Value = "供應商"
    wsQuery.Range("A1:F1").Font.Bold = True

    ' 填入待查詢的商品 ID
    wsQuery.Range("A2").Value = "P003"
    wsQuery.Range("A3").Value = "P007"
    wsQuery.Range("A4").Value = "P001"
    wsQuery.Range("A5").Value = "P005"
    wsQuery.Range("A6").Value = "P009"

    ' 批次填入 INDEX+MATCH 公式（查不到則顯示「查無資料」）
    For i = 2 To 6
        ' 商品名稱（第2欄）
        wsQuery.Cells(i, 2).Formula = "=IFERROR(INDEX('" & refName & "'!$B:$B," & _
            "MATCH(A" & i & ",'" & refName & "'!$A:$A,0)),""查無資料"")"
        ' 類別（第3欄）
        wsQuery.Cells(i, 3).Formula = "=IFERROR(INDEX('" & refName & "'!$C:$C," & _
            "MATCH(A" & i & ",'" & refName & "'!$A:$A,0)),""查無資料"")"
        ' 單價（第4欄）
        wsQuery.Cells(i, 4).Formula = "=IFERROR(INDEX('" & refName & "'!$D:$D," & _
            "MATCH(A" & i & ",'" & refName & "'!$A:$A,0)),0)"
        wsQuery.Cells(i, 4).NumberFormat = "#,##0"
        ' 庫存量（第5欄）
        wsQuery.Cells(i, 5).Formula = "=IFERROR(INDEX('" & refName & "'!$E:$E," & _
            "MATCH(A" & i & ",'" & refName & "'!$A:$A,0)),0)"
        ' 供應商（第6欄）
        wsQuery.Cells(i, 6).Formula = "=IFERROR(INDEX('" & refName & "'!$F:$F," & _
            "MATCH(A" & i & ",'" & refName & "'!$A:$A,0)),""查無資料"")"
    Next i
End Sub

' 填入商品主檔參考資料
Private Sub FillProductMasterData(ByVal ws As Worksheet)
    ws.Range("A1:F1").Value = Array("商品ID", "商品名稱", "類別", "單價", "庫存量", "供應商")
    ws.Range("A1:F1").Font.Bold = True

    Dim data As Variant
    data = Array( _
        Array("P001", "無線滑鼠", "電腦周邊", 590, 120, "科技公司A"), _
        Array("P002", "機械鍵盤", "電腦周邊", 2800, 45, "科技公司B"), _
        Array("P003", "27吋螢幕", "顯示設備", 8500, 30, "顯示公司C"), _
        Array("P004", "USB集線器", "電腦周邊", 450, 200, "科技公司A"), _
        Array("P005", "筆記型電腦", "電腦主機", 35000, 15, "電腦公司D"), _
        Array("P006", "行動硬碟", "儲存設備", 1800, 80, "儲存公司E"), _
        Array("P007", "網路攝影機", "影音設備", 1200, 60, "影音公司F"), _
        Array("P008", "藍芽耳機", "音訊設備", 2200, 90, "音訊公司G"), _
        Array("P009", "桌上型電腦", "電腦主機", 28000, 10, "電腦公司D"), _
        Array("P010", "平板電腦", "行動裝置", 15000, 25, "行動公司H") _
    )

    Dim i As Integer
    For i = 0 To UBound(data)
        ws.Cells(i + 2, 1).Value = data(i)(0)
        ws.Cells(i + 2, 2).Value = data(i)(1)
        ws.Cells(i + 2, 3).Value = data(i)(2)
        ws.Cells(i + 2, 4).Value = data(i)(3)
        ws.Cells(i + 2, 5).Value = data(i)(4)
        ws.Cells(i + 2, 6).Value = data(i)(5)
    Next i

    ws.Range("D2:D11").NumberFormat = "#,##0"
    ws.Columns("A:F").AutoFit
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateIMSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateIMSheet = ws
End Function