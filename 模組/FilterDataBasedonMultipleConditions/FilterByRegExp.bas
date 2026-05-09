Attribute VB_Name = "FilterByRegExp"
Option Explicit

'************************************************************************************
' 模組名稱: FilterByRegExp
' 功能說明: 使用 VBScript.RegExp 對每列進行正則比對，手動控制列可見性
'           適合 AutoFilter 無法表達的複雜文字條件（如：開頭為數字+英文）
'
' 作者版權: Dunk
' 現任設計: Dunk
' 最後修改: 2026/5/9
'************************************************************************************

' 入口：只顯示料號格式符合 "^\d{3}-[A-Z]{2}" 的資料列（如 001-AB）
Public Sub FilterByRegExpPatternExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateWsReg(ThisWorkbook, "RegExp篩選範例")
    Call FillPartNumberData(ws)
    Call ApplyRegExpFilter(ws, 1, "^\d{3}-[A-Z]{2}")

    MsgBox "RegExp 篩選完成！僅顯示符合「3位數字-2位大寫英文」格式的料號。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清除手動隱藏（顯示所有列）
Public Sub ShowAllRows()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Rows.Hidden = False
    MsgBox "已取消所有列隱藏。", vbInformation, "完成"
End Sub

' 對指定欄套用 RegExp 篩選，隱藏不符合的列
Private Sub ApplyRegExpFilter(ByVal ws As Worksheet, _
                               ByVal targetCol As Integer, _
                               ByVal pattern As String)
    Dim re       As Object
    Dim lastRow  As Long
    Dim i        As Long
    Dim cellVal  As String

    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = False
    re.Global = False

    lastRow = ws.Cells(ws.Rows.Count, targetCol).End(xlUp).Row

    ' 先顯示所有列
    ws.Rows.Hidden = False

    ' 從第2列開始逐列比對
    For i = 2 To lastRow
        cellVal = CStr(ws.Cells(i, targetCol).Value)
        If Not re.Test(cellVal) Then
            ws.Rows(i).Hidden = True
        End If
    Next i
End Sub

' 填入料號測試資料
Private Sub FillPartNumberData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("料號", "品名", "庫存量")
    ws.Range("A2:C2").Value = Array("001-AB", "螺絲 M3", 500)
    ws.Range("A3:C3").Value = Array("X12-cd", "螺帽 M4", 300)  ' 小寫不符
    ws.Range("A4:C4").Value = Array("045-ZK", "墊片 10mm", 800)
    ws.Range("A5:C5").Value = Array("AB-123", "彈簧 5cm", 120)  ' 格式不符
    ws.Range("A6:C6").Value = Array("078-TW", "銅管 6分", 200)
    ws.Range("A7:C7").Value = Array("99-ABC", "鋁板 1mm", 150)  ' 格式不符
    ws.Range("A8:C8").Value = Array("123-QR", "不鏽鋼棒", 90)
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表並清空
Private Function GetOrCreateWsReg(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Rows.Hidden = False
    ws.Cells.Clear
    Set GetOrCreateWsReg = ws
End Function