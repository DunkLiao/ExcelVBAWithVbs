Attribute VB_Name = "GradientFillFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: GradientFillFormatting
'功能說明: 依數值大小以三色色階漸層填色方式套用條件式格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub ApplyGradientFillFormatting()
    On Error GoTo ErrHandler
    Dim ws  As Worksheet
    Dim rng As Range

    Set ws = ActiveSheet
    Call FillGradientSampleData(ws)
    Set rng = ws.Range("B2:B11")
    rng.FormatConditions.Delete
    rng.FormatConditions.AddColorScale ColorScaleType:=3

    With rng.FormatConditions(1)
        .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .ColorScaleCriteria(1).FormatColor.Color = RGB(99, 190, 123)
        .ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .ColorScaleCriteria(2).Value = 50
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 132)
        .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .ColorScaleCriteria(3).FormatColor.Color = RGB(248, 105, 107)
    End With
    ws.Columns("A:B").AutoFit
    MsgBox "漸層填色條件格式已套用到 B2:B11 範圍。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillGradientSampleData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "得分"
    ws.Range("A2:B2").Value  = Array("產品 A", 95)
    ws.Range("A3:B3").Value  = Array("產品 B", 42)
    ws.Range("A4:B4").Value  = Array("產品 C", 78)
    ws.Range("A5:B5").Value  = Array("產品 D", 61)
    ws.Range("A6:B6").Value  = Array("產品 E", 88)
    ws.Range("A7:B7").Value  = Array("產品 F", 30)
    ws.Range("A8:B8").Value  = Array("產品 G", 55)
    ws.Range("A9:B9").Value  = Array("產品 H", 73)
    ws.Range("A10:B10").Value = Array("產品 I", 20)
    ws.Range("A11:B11").Value = Array("產品 J", 100)
    With ws.Range("A1:B1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

