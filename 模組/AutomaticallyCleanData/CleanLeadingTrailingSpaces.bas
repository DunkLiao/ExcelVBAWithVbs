Attribute VB_Name = "CleanLeadingTrailingSpaces"
Option Explicit
'*************************************************************************************
'模組名稱: CleanLeadingTrailingSpaces
'功能說明: 自動清除工作表中所有文字儲存格的前置與後置空白字元
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

' 範例進入點
Sub TestCleanLeadingTrailingSpaces()
    Call CleanLeadingTrailingSpacesInSheet
End Sub

' 清除目前工作表或選取範圍中所有文字儲存格的前後空白
Sub CleanLeadingTrailingSpacesInSheet()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim targetRange As Range
    Dim cell As Range
    Dim cleanedCount As Long
    Dim originalValue As String
    Dim cleanedValue As String

    Set ws = ActiveSheet

    If TypeName(Selection) = "Range" And Selection.Cells.Count > 1 Then
        Set targetRange = Selection
    Else
        Set targetRange = ws.UsedRange
    End If

    cleanedCount = 0
    Application.ScreenUpdating = False

    For Each cell In targetRange.Cells
        If cell.HasFormula Then GoTo NextCell
        If Not IsEmpty(cell.Value) Then
            If VarType(cell.Value) = vbString Then
                originalValue = cell.Value
                cleanedValue = Trim(originalValue)
                cleanedValue = Replace(cleanedValue, Chr(12288), "")
                cleanedValue = Trim(cleanedValue)

                If cleanedValue <> originalValue Then
                    cell.Value = cleanedValue
                    cleanedCount = cleanedCount + 1
                End If
            End If
        End If
NextCell:
    Next cell

    Application.ScreenUpdating = True

    MsgBox "前後空白清除完成！" & vbCrLf & _
           "共清理了 " & cleanedCount & " 個儲存格。", _
           vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清除空白時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 建立含前後空白的範例資料並清除
Sub CreateAndCleanSpaceExample()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "空白清除範例")

    ws.Range("A1").Value = "原始資料（含空白）"
    ws.Range("B1").Value = "說明"
    ws.Range("A1:B1").Font.Bold = True

    ws.Range("A2").Value = "  王小明  "
    ws.Range("B2").Value = "前後各兩個空格"

    ws.Range("A3").Value = "  業務部"
    ws.Range("B3").Value = "前置兩個空格"

    ws.Range("A4").Value = "台北市  "
    ws.Range("B4").Value = "後置兩個空格"

    ws.Range("A5").Value = Chr(12288) & "全形空格" & Chr(12288)
    ws.Range("B5").Value = "前後各一個全形空格"

    ws.Range("A6").Value = "   "
    ws.Range("B6").Value = "全部空格（將清空）"

    ws.Columns("A:B").AutoFit
    ws.Range("A2:A6").Select

    Dim answer As Integer
    answer = MsgBox("範例資料已建立完成。" & vbCrLf & _
                    "按「是」立即清除 A2:A6 的前後空白。", _
                    vbYesNo + vbQuestion, "確認清除")

    If answer = vbYes Then
        Call CleanLeadingTrailingSpacesInSheet
    End If

    ws.Activate
    ws.Range("A1").Select
    Exit Sub

ErrorHandler:
    MsgBox "建立範例時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
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
