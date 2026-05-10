Option Explicit
'*************************************************************************************
'模組名稱: EntireRowStatusFormatting
'功能說明: 依狀態欄位套用整列條件格式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Public Sub ApplyEntireRowStatusFormatting()
    Dim ws As Worksheet
    Dim rowRange As Range
    Dim fc As FormatCondition

    On Error GoTo ErrHandler

    Set ws = GetOrCreateStatusSheet("整列狀態格式範例")
    ws.Cells.Clear
    Call FillStatusData(ws)

    Set rowRange = ws.Range("A2:D10")
    rowRange.FormatConditions.Delete

    Set fc = rowRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=$D2=""已完成""")
    With fc
        .Interior.Color = RGB(226, 239, 218)
        .Font.Color = RGB(84, 130, 53)
    End With

    Set fc = rowRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=$D2=""延遲""")
    With fc
        .Interior.Color = RGB(244, 204, 204)
        .Font.Color = RGB(153, 0, 0)
        .Font.Bold = True
    End With

    Set fc = rowRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=$D2=""待處理""")
    With fc
        .Interior.Color = RGB(255, 242, 204)
        .Font.Color = RGB(127, 96, 0)
    End With

    ws.Columns("A:D").AutoFit
    MsgBox "整列狀態條件格式已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立整列狀態條件格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillStatusData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("任務", "負責人", "期限", "狀態")
    ws.Range("A2:D2").Value = Array("盤點庫存", "王小明", Date - 1, "延遲")
    ws.Range("A3:D3").Value = Array("寄送報價", "李小華", Date + 1, "待處理")
    ws.Range("A4:D4").Value = Array("整理合約", "陳美玲", Date - 2, "已完成")
    ws.Range("A5:D5").Value = Array("客戶回訪", "張志強", Date + 4, "進行中")
    ws.Range("A6:D6").Value = Array("發票核對", "林雅婷", Date, "待處理")
    ws.Range("A7:D7").Value = Array("系統測試", "周建宏", Date - 5, "已完成")
    ws.Range("A8:D8").Value = Array("資料備份", "吳佩君", Date - 1, "延遲")
    ws.Range("A9:D9").Value = Array("會議紀錄", "許家豪", Date + 2, "待處理")
    ws.Range("A10:D10").Value = Array("月報確認", "黃怡君", Date + 3, "進行中")
    ws.Range("C2:C10").NumberFormat = "yyyy/m/d"
End Sub

Private Function GetOrCreateStatusSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateStatusSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateStatusSheet Is Nothing Then
        Set GetOrCreateStatusSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateStatusSheet.Name = sheetName
    End If
End Function
