Attribute VB_Name = "StructuredReferenceFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: StructuredReferenceFormulaExample
'功能說明: 示範如何使用 Excel 表格的結構化參照公式（Structured References）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestStructuredReferenceFormula()
    Call CreateStructuredReferenceFormulas
End Sub

Sub CreateStructuredReferenceFormulas()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rngData As Range
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim wsName As String
    wsName = "結構化參照範例"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = wsName
    
    ws.Range("A1").Value = "品名"
    ws.Range("B1").Value = "數量"
    ws.Range("C1").Value = "單價"
    ws.Range("D1").Value = "小計"
    ws.Range("E1").Value = "稅額"
    ws.Range("F1").Value = "總計"
    
    ws.Range("A2").Value = "筆記型電腦"
    ws.Range("B2").Value = 5
    ws.Range("C2").Value = 28000
    ws.Range("A3").Value = "平板電腦"
    ws.Range("B3").Value = 8
    ws.Range("C3").Value = 15000
    ws.Range("A4").Value = "智慧型手機"
    ws.Range("B4").Value = 12
    ws.Range("C4").Value = 22000
    ws.Range("A5").Value = "藍牙耳機"
    ws.Range("B5").Value = 20
    ws.Range("C5").Value = 3500
    ws.Range("A6").Value = "行動電源"
    ws.Range("B6").Value = 15
    ws.Range("C6").Value = 1200
    
    Set rngData = ws.Range("A1:F6")
    
    Set tbl = ws.ListObjects.Add(xlSrcRange, rngData, , xlYes)
    tbl.Name = "銷售資料"
    tbl.TableStyle = "TableStyleMedium2"
    
    ' 使用結構化參照公式
    ws.Range("D2").Formula = "=[@數量]*[@單價]"
    ws.Range("E2").Formula = "=[@小計]*0.05"
    ws.Range("F2").Formula = "=[@小計]+[@稅額]"
    
    ws.Range("A8").Value = "加總結果"
    ws.Range("A8").Font.Bold = True
    ws.Range("A9").Value = "總數量"
    ws.Range("A10").Value = "總金額"
    
    ws.Range("B9").Formula = "=SUM(銷售資料[數量])"
    ws.Range("B10").Formula = "=SUM(銷售資料[總計])"
    ws.Range("B9:B10").NumberFormat = "#,##0"
    
    ws.Columns("A:F").AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "結構化參照公式範例建立完成！" & vbCrLf & _
           "D~F 欄使用 [@欄位名稱] 語法進行結構化參照。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "建立結構化參照公式時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
