Attribute VB_Name = "FilterAndDeleteRows"
Option Explicit
'*************************************************************************************
'模組名稱: FilterAndDeleteRows
'功能說明: 依指定欄位關鍵字或空白值篩選，並刪除符合條件的整列資料
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub DeleteRowsByKeyword()
    On Error GoTo ErrHandler
    Dim ws       As Worksheet
    Dim lastRow  As Long
    Dim colIndex As Long
    Dim keyword  As String
    Dim i        As Long
    Dim delCount As Long
    Dim cellVal  As String

    Set ws = ActiveSheet
    keyword  = InputBox("請輸入要刪除的關鍵字（包含此文字的整列將被刪除）：", "刪除列條件")
    If keyword = "" Then
        MsgBox "未輸入關鍵字，已取消操作。", vbExclamation, "提示"
        Exit Sub
    End If
    colIndex = CInt(InputBox("請輸入要比對的欄號（例如：1 表示第 A 欄）：", "欄號", "1"))
    If colIndex < 1 Then
        MsgBox "欄號無效，已取消操作。", vbExclamation, "提示"
        Exit Sub
    End If
    lastRow  = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
    delCount = 0
    For i = lastRow To 2 Step -1
        cellVal = CStr(ws.Cells(i, colIndex).Value)
        If InStr(1, cellVal, keyword, vbTextCompare) > 0 Then
            ws.Rows(i).Delete
            delCount = delCount + 1
        End If
    Next i
    MsgBox "已刪除 " & delCount & " 列包含「" & keyword & "」的資料。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Public Sub DeleteRowsByBlankColumn()
    On Error GoTo ErrHandler
    Dim ws       As Worksheet
    Dim lastRow  As Long
    Dim colIndex As Long
    Dim i        As Long
    Dim delCount As Long

    Set ws = ActiveSheet
    colIndex = CInt(InputBox("請輸入要檢查空白的欄號（例如：1 表示 A 欄）：", "欄號", "1"))
    If colIndex < 1 Then
        MsgBox "欄號無效，已取消操作。", vbExclamation, "提示"
        Exit Sub
    End If
    lastRow  = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row
    delCount = 0
    For i = lastRow To 2 Step -1
        If Trim(CStr(ws.Cells(i, colIndex).Value)) = "" Then
            ws.Rows(i).Delete
            delCount = delCount + 1
        End If
    Next i
    MsgBox "已刪除 " & delCount & " 列空白列。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Public Sub CreateFilterDeleteSampleData()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("篩選刪除範例")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "篩選刪除範例"
    End If
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("編號", "狀態", "金額")
    ws.Range("A2:C2").Value = Array(1, "已完成", 5000)
    ws.Range("A3:C3").Value = Array(2, "已取消", 3000)
    ws.Range("A4:C4").Value = Array(3, "已完成", 8000)
    ws.Range("A5:C5").Value = Array(4, "已取消", 2000)
    ws.Range("A6:C6").Value = Array(5, "處理中", 6000)
    ws.Range("A7:C7").Value = Array(6, "已取消", 1500)
    ws.Range("A1:C1").Font.Bold = True
    ws.Columns.AutoFit
    ws.Activate
    MsgBox "測試資料已建立。執行 DeleteRowsByKeyword，輸入「已取消」即可刪除對應列。", vbInformation, "完成"
End Sub

