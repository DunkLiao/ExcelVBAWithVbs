Attribute VB_Name = "InventoryLevelFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: InventoryLevelFormatting
'功能說明: 依庫存量設定條件式格式：
'          低庫存（< 安全庫存）標紅；正常庫存標綠；超量庫存標黃
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub ApplyInventoryLevelFormatting()
    Dim ws          As Worksheet
    Dim dataRange   As Range
    Dim lastRow     As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("庫存水位範例")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "庫存水位範例"
    End If

    ws.Cells.Clear
    Call FillInventoryData(ws)

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Set dataRange = ws.Range("C2:C" & lastRow)

    dataRange.FormatConditions.Delete

    ' 低庫存：庫存量 < 安全庫存（B欄）→ 紅色背景
    Dim fc1 As FormatCondition
    Set fc1 = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=C2<B2")
    fc1.Interior.Color = RGB(255, 199, 206)
    fc1.Font.Color = RGB(156, 0, 6)
    fc1.Font.Bold = True

    ' 超量庫存：庫存量 > 安全庫存 * 3 → 黃色背景
    Dim fc2 As FormatCondition
    Set fc2 = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=C2>B2*3")
    fc2.Interior.Color = RGB(255, 235, 156)
    fc2.Font.Color = RGB(156, 87, 0)

    ' 正常庫存：庫存量 >= 安全庫存 且 <= 安全庫存 * 3 → 綠色背景
    Dim fc3 As FormatCondition
    Set fc3 = dataRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=AND(C2>=B2,C2<=B2*3)")
    fc3.Interior.Color = RGB(198, 239, 206)
    fc3.Font.Color = RGB(0, 97, 0)

    ws.Columns.AutoFit
    MsgBox "庫存水位條件格式已套用完成！" & Chr(10) & _
        "紅色 = 低庫存（低於安全庫存）" & Chr(10) & _
        "綠色 = 正常庫存" & Chr(10) & _
        "黃色 = 超量庫存（超過安全庫存 3 倍）", vbInformation, "完成"
End Sub

Private Sub FillInventoryData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "商品名稱"
    ws.Range("B1").Value = "安全庫存"
    ws.Range("C1").Value = "現有庫存"
    ws.Range("D1").Value = "狀態說明"

    Dim data(5, 3) As Variant
    data(0, 0) = "產品A" : data(0, 1) = 50  : data(0, 2) = 20  : data(0, 3) = "低庫存"
    data(1, 0) = "產品B" : data(1, 1) = 30  : data(1, 2) = 80  : data(1, 3) = "正常"
    data(2, 0) = "產品C" : data(2, 1) = 100 : data(2, 2) = 350 : data(2, 3) = "超量"
    data(3, 0) = "產品D" : data(3, 1) = 200 : data(3, 2) = 150 : data(3, 3) = "低庫存"
    data(4, 0) = "產品E" : data(4, 1) = 80  : data(4, 2) = 90  : data(4, 3) = "正常"
    data(5, 0) = "產品F" : data(5, 1) = 60  : data(5, 2) = 250 : data(5, 3) = "超量"

    Dim i As Integer
    For i = 0 To 5
        ws.Cells(i + 2, 1).Value = data(i, 0)
        ws.Cells(i + 2, 2).Value = data(i, 1)
        ws.Cells(i + 2, 3).Value = data(i, 2)
        ws.Cells(i + 2, 4).Value = data(i, 3)
    Next i

    ws.Range("A1:D1").Font.Bold = True
End Sub
