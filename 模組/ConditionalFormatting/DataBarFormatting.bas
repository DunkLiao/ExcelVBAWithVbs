Attribute VB_Name = "DataBarFormatting"
Option Explicit

' ============================================================
' 範例：以 VBA 設定資料橫條（Data Bar）條件式格式
' 功能：對選取範圍套用資料橫條，並自訂顏色與最大值設定
' ============================================================
Sub ApplyDataBarFormatting()
    Dim ws      As Worksheet
    Dim rng     As Range
    Dim db      As Databar

    On Error GoTo ErrHandler

    ' --- 準備示範資料 ---
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "DataBarDemo"

    ws.Cells(1, 1).Value = "產品"
    ws.Cells(1, 2).Value = "銷售量"
    ws.Cells(2, 1).Value = "產品A"
    ws.Cells(2, 2).Value = 320
    ws.Cells(3, 1).Value = "產品B"
    ws.Cells(3, 2).Value = 150
    ws.Cells(4, 1).Value = "產品C"
    ws.Cells(4, 2).Value = 480
    ws.Cells(5, 1).Value = "產品D"
    ws.Cells(5, 2).Value = 210
    ws.Cells(6, 1).Value = "產品E"
    ws.Cells(6, 2).Value = 390

    Set rng = ws.Range("B2:B6")

    ' --- 清除既有條件式格式 ---
    rng.FormatConditions.Delete

    ' --- 新增資料橫條格式 ---
    Set db = rng.FormatConditions.AddDatabar

    ' --- 設定最小值為 0，最大值自動 ---
    db.MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
    db.MaxPoint.Modify newtype:=xlConditionValueAutomaticMax

    ' --- 設定橫條顏色（藍色）---
    db.BarColor.Color = RGB(0, 112, 192)
    db.BarFillType = xlDataBarFillGradient

    ' --- 顯示數值 ---
    db.ShowValue = True

    ws.Columns.AutoFit
    MsgBox "資料橫條條件式格式已套用至範圍：" & rng.Address, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
