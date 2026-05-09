Attribute VB_Name = "FilterByDynamicInputBox"
Option Explicit

'************************************************************************************
' 模組名稱: FilterByDynamicInputBox
' 功能說明: 透過 InputBox 讓使用者即時輸入篩選條件（欄位名稱＋條件值）
'           動態決定要篩選哪一欄與比較值，實現互動式篩選
'
' 作者版權: Dunk
' 現任設計: Dunk
' 最後修改: 2026/5/9
'************************************************************************************

' 入口：提示使用者輸入條件後執行篩選
Public Sub FilterByDynamicInputBoxExample()
    On Error GoTo ErrHandler

    Dim ws         As Worksheet
    Dim colName    As String
    Dim criteria   As String
    Dim fieldIndex As Integer

    Set ws = GetOrCreateWsDyn(ThisWorkbook, "動態篩選範例")
    Call FillInventoryData(ws)

    ' 詢問要篩選的欄位名稱
    colName = InputBox("請輸入篩選欄位名稱（倉庫 / 類別 / 數量）：", "動態篩選 - 欄位", "倉庫")
    If colName = "" Then Exit Sub

    ' 詢問條件值
    criteria = InputBox("請輸入篩選條件（可加比較符號如 >100）：", "動態篩選 - 條件", "A倉")
    If criteria = "" Then Exit Sub

    ' 查找欄位編號
    fieldIndex = FindColumnIndex(ws, colName)
    If fieldIndex = 0 Then
        MsgBox "找不到欄位「" & colName & "」，請確認欄名正確。", vbExclamation, "欄位錯誤"
        Exit Sub
    End If

    ' 套用篩選
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1").CurrentRegion.AutoFilter Field:=fieldIndex, Criteria1:=criteria
    ws.Columns("A:D").AutoFit

    MsgBox "已篩選「" & colName & "」" & criteria & " 的資料。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清除動態篩選
Public Sub ClearDynamicFilter()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    MsgBox "已清除篩選。", vbInformation, "完成"
End Sub

' 依標題名稱找欄位編號（1-based）
Private Function FindColumnIndex(ByVal ws As Worksheet, ByVal colName As String) As Integer
    Dim lastCol As Integer
    Dim j       As Integer

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    FindColumnIndex = 0

    For j = 1 To lastCol
        If ws.Cells(1, j).Value = colName Then
            FindColumnIndex = j
            Exit Function
        End If
    Next j
End Function

' 填入倉庫庫存測試資料
Private Sub FillInventoryData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("料號", "類別", "倉庫", "數量")
    ws.Range("A2:D2").Value = Array("P001", "零件", "A倉", 250)
    ws.Range("A3:D3").Value = Array("P002", "成品", "B倉", 80)
    ws.Range("A4:D4").Value = Array("P003", "零件", "A倉", 420)
    ws.Range("A5:D5").Value = Array("P004", "原料", "C倉", 1500)
    ws.Range("A6:D6").Value = Array("P005", "成品", "A倉", 35)
    ws.Range("A7:D7").Value = Array("P006", "原料", "B倉", 200)
    ws.Range("A8:D8").Value = Array("P007", "零件", "C倉", 90)
    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表並清空
Private Function GetOrCreateWsDyn(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateWsDyn = ws
End Function