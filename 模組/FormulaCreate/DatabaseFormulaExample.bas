Attribute VB_Name = "DatabaseFormulaExample"
Option Explicit

' ============================================================
' 範例：使用 VBA 建立資料庫函數公式（DSUM / DAVERAGE / DCOUNT）
' 功能：示範在工作表中插入資料庫函數以進行條件彙總
' ============================================================
Sub CreateDatabaseFormulaExample()
    Dim ws      As Worksheet
    Dim i       As Integer
    Dim arrData As Variant

    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "DBFormulaDemo"

    ' --- 建立資料庫標題列 ---
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "員工"
    ws.Range("C1").Value = "薪資"
    ws.Range("D1").Value = "年資"

    ' --- 填入示範資料 ---
    arrData = Array( _
        Array("業務", "王小明", 45000, 3), _
        Array("業務", "林美華", 42000, 2), _
        Array("工程", "張大同", 60000, 5), _
        Array("工程", "陳雅婷", 58000, 4), _
        Array("人事", "劉建國", 40000, 1))
    For i = 0 To 4
        ws.Cells(2 + i, 1).Value = arrData(i)(0)
        ws.Cells(2 + i, 2).Value = arrData(i)(1)
        ws.Cells(2 + i, 3).Value = arrData(i)(2)
        ws.Cells(2 + i, 4).Value = arrData(i)(3)
    Next i

    ' --- 建立條件區域 ---
    ws.Range("F1").Value = "部門"
    ws.Range("F2").Value = "業務"

    ' --- 插入資料庫函數 ---
    ws.Range("H1").Value = "業務部薪資合計"
    ws.Range("H2").Formula = "=DSUM(A1:D6,""薪資"",F1:F2)"

    ws.Range("H4").Value = "業務部平均薪資"
    ws.Range("H5").Formula = "=DAVERAGE(A1:D6,""薪資"",F1:F2)"

    ws.Range("H7").Value = "業務部人數"
    ws.Range("H8").Formula = "=DCOUNT(A1:D6,""薪資"",F1:F2)"

    ws.Columns.AutoFit
    MsgBox "資料庫函數公式已建立於工作表：" & ws.Name, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
