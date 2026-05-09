Attribute VB_Name = "FilterBlankNonBlank"
Option Explicit

'************************************************************************************
' 模組名稱: FilterBlankNonBlank
' 功能說明: 使用 AutoFilter 篩選空白列或非空白列
'           示範找出備註未填寫與已填寫的記錄
'
' 作者版權: Dunk
' 現任設計: Dunk
' 最後修改: 2026/5/9
'************************************************************************************

' 入口：篩選備註欄空白的列（待跟進記錄）
Public Sub FilterBlankRemarkExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateWsBlank(ThisWorkbook, "空白篩選範例")
    Call FillRemarkData(ws)
    Call ApplyBlankFilter(ws, 4, True)

    MsgBox "已篩選出「備註」欄空白的記錄（待跟進）。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 入口：篩選備註欄非空白的列（已處理記錄）
Public Sub FilterNonBlankRemarkExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateWsBlank(ThisWorkbook, "空白篩選範例")
    Call FillRemarkData(ws)
    Call ApplyBlankFilter(ws, 4, False)

    MsgBox "已篩選出「備註」欄有填寫的記錄（已處理）。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清除篩選
Public Sub ClearBlankFilter()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    MsgBox "已清除篩選條件。", vbInformation, "完成"
End Sub

' 套用空白/非空白篩選
' filterField: 欄位編號（1-based）
' blankOnly: True=只顯示空白，False=只顯示非空白
Private Sub ApplyBlankFilter(ByVal ws As Worksheet, _
                              ByVal filterField As Integer, _
                              ByVal blankOnly As Boolean)
    Dim rng As Range
    Set rng = ws.Range("A1").CurrentRegion
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    If blankOnly Then
        rng.AutoFilter Field:=filterField, Criteria1:="="
    Else
        rng.AutoFilter Field:=filterField, Criteria1:="<>"
    End If
    ws.Columns("A:D").AutoFit
End Sub

' 填入含空白備註的測試資料
Private Sub FillRemarkData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("案件編號", "客戶", "金額", "備註")
    ws.Range("A2:D2").Value = Array("C001", "甲公司", 15000, "已確認")
    ws.Range("A3:D3").Value = Array("C002", "乙公司", 8500, "")
    ws.Range("A4:D4").Value = Array("C003", "丙公司", 22000, "待簽約")
    ws.Range("A5:D5").Value = Array("C004", "丁公司", 3200, "")
    ws.Range("A6:D6").Value = Array("C005", "戊公司", 67000, "已付款")
    ws.Range("A7:D7").Value = Array("C006", "己公司", 4800, "")
    ws.Range("C2:C7").NumberFormat = "#,##0"
    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表並清空
Private Function GetOrCreateWsBlank(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateWsBlank = ws
End Function