Attribute VB_Name = "DynamicRangeFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: DynamicRangeFormatting
'功能說明: 設定條件式格式時使用動態範圍，當資料增減時格式自動調整
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestDynamicRangeFormatting()
    Call ApplyDynamicRangeFormatting
End Sub

Sub ApplyDynamicRangeFormatting()
    Dim ws As Worksheet
    Dim dynamicRange As Range
    Dim cf As FormatCondition
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim wsName As String
    wsName = "動態範圍格式"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = wsName
    
    ' 撰寫範例資料
    ws.Range("A1").Value = "產品名稱"
    ws.Range("B1").Value = "庫存量"
    ws.Range("C1").Value = "安全庫存"
    ws.Range("D1").Value = "狀態"
    
    Dim products As Variant
    products = Array("產品A", "產品B", "產品C", "產品D", "產品E", "產品F", "產品G", "產品H")
    Dim stocks As Variant
    stocks = Array(120, 45, 200, 30, 150, 80, 10, 95)
    Dim safeStocks As Variant
    safeStocks = Array(50, 50, 50, 50, 50, 50, 50, 50)
    
    Dim i As Long
    For i = 1 To 8
        ws.Cells(i + 1, 1).Value = products(i - 1)
        ws.Cells(i + 1, 2).Value = stocks(i - 1)
        ws.Cells(i + 1, 3).Value = safeStocks(i - 1)
    Next i
    
    For i = 1 To 8
        If ws.Cells(i + 1, 2).Value < ws.Cells(i + 1, 3).Value Then
            ws.Cells(i + 1, 4).Value = "庫存不足"
        ElseIf ws.Cells(i + 1, 2).Value < ws.Cells(i + 1, 3).Value * 2 Then
            ws.Cells(i + 1, 4).Value = "注意補貨"
        Else
            ws.Cells(i + 1, 4).Value = "庫存充足"
        End If
    Next i
    
    ' 使用動態命名範圍
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set dynamicRange = ws.Range("B2:B" & lastRow)
    
    ' 清除現有條件式格式
    dynamicRange.FormatConditions.Delete
    
    ' 條件1：庫存低於安全庫存 - 紅色
    Set cf = dynamicRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlLess, _
        Formula1:="=C2")
    cf.Interior.Color = RGB(255, 199, 206)
    cf.Font.Color = RGB(156, 0, 6)
    
    ' 條件2：庫存介於安全庫存與2倍安全庫存之間 - 黃色
    Set cf = dynamicRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlBetween, _
        Formula1:="=C2", _
        Formula2:="=C2*2")
    cf.Interior.Color = RGB(255, 235, 156)
    cf.Font.Color = RGB(156, 101, 0)
    
    ' 條件3：庫存高於2倍安全庫存 - 綠色
    Set cf = dynamicRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlGreater, _
        Formula1:="=C2*2")
    cf.Interior.Color = RGB(198, 239, 206)
    cf.Font.Color = RGB(0, 97, 0)
    
    ' 整列格式
    Dim rngAll As Range
    Set rngAll = ws.Range("A2:D" & lastRow)
    rngAll.FormatConditions.Delete
    
    Set cf = rngAll.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=$D2=""庫存不足""")
    cf.Interior.Color = RGB(255, 199, 206)
    
    Set cf = rngAll.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=$D2=""注意補貨""")
    cf.Interior.Color = RGB(255, 235, 156)
    
    Set cf = rngAll.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=$D2=""庫存充足""")
    cf.Interior.Color = RGB(198, 239, 206)
    
    ws.Columns.AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "動態範圍條件式格式設定完成！" & vbCrLf & _
           "新增或刪除資料列後，格式範圍會自動調整。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "設定條件式格式時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
