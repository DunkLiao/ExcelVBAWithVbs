Attribute VB_Name = "AdvancedFilterExample"
Option Explicit

' ============================================================
' 範例：使用 VBA 進階篩選（AdvancedFilter）篩選資料至新位置
' 功能：設定條件範圍後，將篩選結果複製到同工作表指定區域
' ============================================================
Sub RunAdvancedFilterExample()
    Dim ws          As Worksheet
    Dim rngData     As Range
    Dim rngCriteria As Range
    Dim rngOutput   As Range

    On Error GoTo ErrHandler

    ' --- 建立示範資料 ---
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "AdvFilterDemo"

    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "員工"
    ws.Range("C1").Value = "薪資"

    Dim arrData As Variant
    arrData = Array( _
        Array("業務", "王小明", 45000), _
        Array("業務", "林美華", 42000), _
        Array("工程", "張大同", 60000), _
        Array("工程", "陳雅婷", 58000), _
        Array("人事", "劉建國", 40000), _
        Array("業務", "吳志明", 48000))
    Dim i As Integer
    For i = 0 To 5
        ws.Cells(2 + i, 1).Value = arrData(i)(0)
        ws.Cells(2 + i, 2).Value = arrData(i)(1)
        ws.Cells(2 + i, 3).Value = arrData(i)(2)
    Next i

    ' --- 設定條件範圍（部門=業務，薪資>42000）---
    ws.Range("E1").Value = "部門"
    ws.Range("E2").Value = "業務"
    ws.Range("F1").Value = "薪資"
    ws.Range("F2").Value = ">42000"

    ' --- 設定輸出標題 ---
    ws.Range("A10").Value = "篩選結果："

    Set rngData = ws.Range("A1:C7")
    Set rngCriteria = ws.Range("E1:F2")
    Set rngOutput = ws.Range("A11")

    ' --- 執行進階篩選 ---
    rngData.AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=rngCriteria, _
        CopyToRange:=rngOutput, _
        Unique:=False

    ws.Columns.AutoFit
    MsgBox "進階篩選完成！結果已複製至 A11 起始範圍。", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
