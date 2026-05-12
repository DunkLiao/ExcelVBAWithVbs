Attribute VB_Name = "MultiColumnRuleFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: MultiColumnRuleFormatting
'功能說明: 對多個欄位同時套用獨立的條件式格式規則，提升多欄資料可讀性
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestMultiColumnRuleFormatting()
    Call ApplyMultiColumnRuleFormatting
End Sub

' 對多欄套用條件式格式
Sub ApplyMultiColumnRuleFormatting()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Set ws = GetOrCreateSheet(ThisWorkbook, "多欄條件格式")

    Call FillMultiColumnData(ws)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 清除原有條件式格式
    ws.Cells.FormatConditions.Delete

    ' 欄B（銷售額）：大於 15000 標綠色
    With ws.Range("B2:B" & lastRow).FormatConditions.Add( _
        Type:=xlCellValue, Operator:=xlGreater, Formula1:="15000")
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With

    ' 欄B（銷售額）：小於 10000 標紅色
    With ws.Range("B2:B" & lastRow).FormatConditions.Add( _
        Type:=xlCellValue, Operator:=xlLess, Formula1:="10000")
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With

    ' 欄C（達成率）：大於等於 100% 標藍色加粗
    With ws.Range("C2:C" & lastRow).FormatConditions.Add( _
        Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="1")
        .Interior.Color = RGB(189, 215, 238)
        .Font.Color = RGB(31, 73, 125)
        .Font.Bold = True
    End With

    ' 欄D（評等）：包含「優」標金色
    With ws.Range("D2:D" & lastRow).FormatConditions.Add( _
        Type:=xlTextString, String:="優", TextOperator:=xlContains)
        .Interior.Color = RGB(255, 235, 156)
        .Font.Color = RGB(156, 101, 0)
    End With

    ws.Columns("A:D").AutoFit

    MsgBox "多欄條件式格式已套用完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "套用條件式格式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 填入多欄範例資料
Private Sub FillMultiColumnData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "業務員"
    ws.Range("B1").Value = "銷售額"
    ws.Range("C1").Value = "達成率"
    ws.Range("D1").Value = "評等"
    ws.Range("A1:D1").Font.Bold = True

    ws.Range("A2").Value = "張小明"
    ws.Range("B2").Value = 18000
    ws.Range("C2").Value = 1.2
    ws.Range("D2").Value = "優"

    ws.Range("A3").Value = "李美玲"
    ws.Range("B3").Value = 8500
    ws.Range("C3").Value = 0.57
    ws.Range("D3").Value = "差"

    ws.Range("A4").Value = "王大華"
    ws.Range("B4").Value = 13500
    ws.Range("C4").Value = 0.9
    ws.Range("D4").Value = "良"

    ws.Range("A5").Value = "陳俊宏"
    ws.Range("B5").Value = 16200
    ws.Range("C5").Value = 1.08
    ws.Range("D5").Value = "優"

    ws.Range("A6").Value = "林淑芬"
    ws.Range("B6").Value = 9200
    ws.Range("C6").Value = 0.61
    ws.Range("D6").Value = "差"

    ws.Range("C2:C6").NumberFormat = "0%"
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
