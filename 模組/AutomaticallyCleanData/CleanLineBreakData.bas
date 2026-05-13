Attribute VB_Name = "CleanLineBreakData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanLineBreakData
'功能說明: 清除儲存格中的換行符號（Chr(10)/Chr(13)），並壓縮多餘空白
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub CleanLineBreaksInSelection()
    On Error GoTo ErrHandler
    Dim rng      As Range
    Dim cell     As Range
    Dim original As String
    Dim cleaned  As String
    Dim count    As Long
    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清除換行的儲存格範圍。", vbExclamation, "提示"
        Exit Sub
    End If
    Set rng = Selection
    count = 0
    For Each cell In rng.Cells
        If cell.Value <> "" And Not IsError(cell.Value) Then
            original = CStr(cell.Value)
            cleaned  = original
            cleaned  = Replace(cleaned, Chr(10), " ")
            cleaned  = Replace(cleaned, Chr(13), " ")
            Do While InStr(cleaned, "  ") > 0
                cleaned = Replace(cleaned, "  ", " ")
            Loop
            cleaned = Trim(cleaned)
            If cleaned <> original Then
                cell.Value    = cleaned
                cell.WrapText = False
                count = count + 1
            End If
        End If
    Next cell
    MsgBox "已清除換行符號，共處理 " & count & " 個儲存格。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Public Sub CleanLineBreaksInActiveSheet()
    On Error GoTo ErrHandler
    Dim ws    As Worksheet
    Dim rng   As Range
    Dim cell  As Range
    Dim val   As String
    Dim count As Long
    Set ws  = ActiveSheet
    Set rng = ws.UsedRange
    count   = 0
    For Each cell In rng.Cells
        If cell.Value <> "" And Not IsError(cell.Value) Then
            val = CStr(cell.Value)
            If InStr(val, Chr(10)) > 0 Or InStr(val, Chr(13)) > 0 Then
                val = Replace(val, Chr(10), " ")
                val = Replace(val, Chr(13), " ")
                Do While InStr(val, "  ") > 0
                    val = Replace(val, "  ", " ")
                Loop
                cell.Value    = Trim(val)
                cell.WrapText = False
                count = count + 1
            End If
        End If
    Next cell
    MsgBox "工作表「" & ws.Name & "」已清除換行符號，共處理 " & count & " 個儲存格。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Public Sub CreateLineBreakSampleData()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("換行清除範例")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "換行清除範例"
    End If
    ws.Cells.Clear
    ws.Range("A1").Value = "姓名"
    ws.Range("B1").Value = "備註（含換行）"
    ws.Range("A2").Value = "張小明"
    ws.Range("B2").Value = "業務部" & Chr(10) & "台北辦公室"
    ws.Range("A3").Value = "李美華"
    ws.Range("B3").Value = "財務部" & Chr(13) & Chr(10) & "高雄辦公室"
    ws.Range("A4").Value = "王大同"
    ws.Range("B4").Value = "研發部" & Chr(10) & "新竹辦公室" & Chr(10) & "遠端上班"
    ws.Range("B2:B4").WrapText = True
    ws.Range("A1:B1").Font.Bold = True
    ws.Columns("A:B").AutoFit
    ws.Rows("2:4").RowHeight = 50
    ws.Activate
    MsgBox "換行範例資料已建立，請執行 CleanLineBreaksInActiveSheet 清除換行。", vbInformation, "完成"
End Sub

