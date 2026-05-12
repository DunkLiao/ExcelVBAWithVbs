Attribute VB_Name = "CleanInvisibleCharacters"
Option Explicit
'*************************************************************************************
'模組名稱: CleanInvisibleCharacters
'功能說明: 自動清除工作表儲存格中的不可見控制字元（ASCII 0-31），保留正常文字內容
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestCleanInvisibleCharacters()
    Call CleanInvisibleCharactersInSheet
End Sub

' 清除現用工作表所有儲存格中的不可見字元
Sub CleanInvisibleCharactersInSheet()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim targetRange As Range
    Dim cell As Range
    Dim originalValue As String
    Dim cleanedValue As String
    Dim cleanedCount As Long
    Dim charCode As Integer

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
                cleanedValue = originalValue

                ' 移除 ASCII 0~31 的控制字元（保留 Tab=9、換行=10、CR=13）
                For charCode = 0 To 31
                    If charCode <> 9 And charCode <> 10 And charCode <> 13 Then
                        cleanedValue = Replace(cleanedValue, Chr(charCode), "")
                    End If
                Next charCode

                ' 移除 ASCII 127 (DEL)
                cleanedValue = Replace(cleanedValue, Chr(127), "")

                If cleanedValue <> originalValue Then
                    cell.Value = cleanedValue
                    cleanedCount = cleanedCount + 1
                End If
            End If
        End If
NextCell:
    Next cell

    Application.ScreenUpdating = True

    MsgBox "不可見字元清除完成！" & vbCrLf & _
           "共清理了 " & cleanedCount & " 個儲存格。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清除不可見字元時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 建立含不可見字元的範例資料並執行清理
Sub CreateAndCleanInvisibleExample()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim answer As Integer
    Set ws = GetOrCreateSheet(ThisWorkbook, "不可見字元清理")

    ws.Range("A1").Value = "原始資料（含不可見字元）"
    ws.Range("B1").Value = "說明"
    ws.Range("A1:B1").Font.Bold = True

    ws.Range("A2").Value = "正常文字" & Chr(0) & "含NUL字元"
    ws.Range("B2").Value = "含 ASCII 0 (NUL)"

    ws.Range("A3").Value = Chr(1) & "開頭有SOH字元"
    ws.Range("B3").Value = "含 ASCII 1 (SOH)"

    ws.Range("A4").Value = "中間" & Chr(7) & "有BEL字元"
    ws.Range("B4").Value = "含 ASCII 7 (BEL)"

    ws.Range("A5").Value = "結尾有DEL" & Chr(127)
    ws.Range("B5").Value = "含 ASCII 127 (DEL)"

    ws.Range("A6").Value = "純淨文字，無控制字元"
    ws.Range("B6").Value = "正常資料"

    ws.Columns("A:B").AutoFit
    ws.Range("A2:A6").Select

    answer = MsgBox("範例資料已建立（含不可見字元）。" & vbCrLf & _
                    "按是立即清除 A2:A6 的不可見字元。", _
                    vbYesNo + vbQuestion, "確認清除")

    If answer = vbYes Then
        Call CleanInvisibleCharactersInSheet
    End If

    ws.Activate
    ws.Range("A1").Select
    Exit Sub

ErrorHandler:
    MsgBox "建立範例時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
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
