Attribute VB_Name = "CrossColumnFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: CrossColumnFormatting
'功能說明: 依據 B 欄數值跨欄位套用條件式格式並醒目顯示整列資料
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

Public Sub RunCrossColumnFormatting()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim targetRange As Range
    Dim formatRule As FormatCondition

    Set ws = GetOrCreateCrossFormatSheet("跨欄位條件格式")
    ws.Cells.Clear
    Call FillCrossFormatData(ws)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set targetRange = ws.Range("A2:C" & lastRow)
    targetRange.FormatConditions.Delete

    Set formatRule = targetRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=$B2>100")

    With formatRule
        .Interior.Color = RGB(255, 255, 0)
        .Font.Bold = True
    End With

    ws.Columns("A:C").AutoFit
    MsgBox "跨欄位條件式格式已套用完成。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "套用條件式格式時發生錯誤: " & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillCrossFormatData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("品項", "數量", "備註")
    ws.Range("A2:C2").Value = Array("滑鼠", 88, "正常")
    ws.Range("A3:C3").Value = Array("鍵盤", 105, "需補貨")
    ws.Range("A4:C4").Value = Array("螢幕", 120, "優先處理")
    ws.Range("A5:C5").Value = Array("主機", 64, "正常")
    ws.Range("A6:C6").Value = Array("印表機", 145, "高需求")
End Sub

Private Function GetOrCreateCrossFormatSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateCrossFormatSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateCrossFormatSheet Is Nothing Then
        Set GetOrCreateCrossFormatSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateCrossFormatSheet.Name = sheetName
    End If
End Function
