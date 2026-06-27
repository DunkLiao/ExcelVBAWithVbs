Attribute VB_Name = "BlinkAlertConditionalFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: BlinkAlertConditionalFormatting
'功能說明: 針對低於安全庫存的項目，使用條件式格式以紅黃綠三色燈號警示庫存狀態
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestBlinkAlertConditionalFormatting()
    Call SetupBlinkAlert
End Sub

Sub SetupBlinkAlert()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    
    sheetName = "庫存警示"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillInventoryData(ws)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ws.Cells.FormatConditions.Delete
    
    With ws.Range("A2:D" & lastRow)
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=$D2<$C2*0.2"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 100, 100)
        .FormatConditions(.FormatConditions.Count).Font.Color = RGB(255, 255, 255)
        .FormatConditions(.FormatConditions.Count).Font.Bold = True
    End With
    
    With ws.Range("A2:D" & lastRow)
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=AND($D2>=$C2*0.2,$D2<$C2*0.5)"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 255, 150)
        .FormatConditions(.FormatConditions.Count).Font.Bold = True
    End With
    
    With ws.Range("A2:D" & lastRow)
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=$D2>=$C2*0.5"
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(200, 255, 200)
    End With
    
    ws.Columns("A:D").AutoFit
    ws.Activate
    
    MsgBox "庫存警示條件式格式已設定完成！" & vbCrLf & _
           "紅色 = 低於安全庫存20%（緊急）" & vbCrLf & _
           "黃色 = 低於安全庫存50%（注意）" & vbCrLf & _
           "綠色 = 庫存充足", vbInformation, "完成"
End Sub

Private Sub FillInventoryData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品編號"
    ws.Range("B1").Value = "產品名稱"
    ws.Range("C1").Value = "安全庫存"
    ws.Range("D1").Value = "目前庫存"
    
    ws.Range("A2").Value = "M001"
    ws.Range("B2").Value = "螺絲"
    ws.Range("C2").Value = 1000
    ws.Range("D2").Value = 150
    
    ws.Range("A3").Value = "M002"
    ws.Range("B3").Value = "墊片"
    ws.Range("C3").Value = 500
    ws.Range("D3").Value = 200
    
    ws.Range("A4").Value = "M003"
    ws.Range("B4").Value = "彈簧"
    ws.Range("C4").Value = 800
    ws.Range("D4").Value = 600
    
    ws.Range("A5").Value = "M004"
    ws.Range("B5").Value = "軸承"
    ws.Range("C5").Value = 300
    ws.Range("D5").Value = 50
    
    ws.Range("A6").Value = "M005"
    ws.Range("B6").Value = "齒輪"
    ws.Range("C6").Value = 200
    ws.Range("D6").Value = 180
    
    ws.Range("A7").Value = "M006"
    ws.Range("B7").Value = "皮帶"
    ws.Range("C7").Value = 400
    ws.Range("D7").Value = 350
    
    ws.Columns("A:D").AutoFit
End Sub
