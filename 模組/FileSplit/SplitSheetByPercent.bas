Option Explicit
Attribute VB_Name = "SplitSheetByPercent"
'*************************************************************************************
'模組名稱: SplitSheetByPercent
'功能說明: 依指定百分比拆分工作表資料到多個新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestSplitByPercent()
    Call SplitSheetByPercentage
End Sub

' 依百分比拆分工作表資料
Sub SplitSheetByPercentage()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalRows As Long
    Dim percent As Double
    Dim splitCount As Long
    Dim wsNew As Worksheet
    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim percentStr As String

    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "請先選取一個工作表。", vbExclamation, "提示"
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "工作表沒有足夠的資料可供拆分。", vbExclamation, "提示"
        Exit Sub
    End If

    totalRows = lastRow - 1

    percentStr = InputBox("請輸入每個區塊的百分比（如 25 代表 25%）：", "拆分百分比", "25")
    If percentStr = "" Then Exit Sub
    If Not IsNumeric(percentStr) Then
        MsgBox "請輸入有效的數字。", vbExclamation, "提示"
        Exit Sub
    End If
    percent = CDbl(percentStr) / 100

    If percent <= 0 Or percent > 1 Then
        MsgBox "百分比必須介於 1 到 100 之間。", vbExclamation, "提示"
        Exit Sub
    End If

    ' 計算每個區塊的列數
    splitCount = WorksheetFunction.RoundUp(totalRows * percent, 0)
    If splitCount < 1 Then splitCount = 1

    ' 開始拆分
    startRow = 2
    i = 1
    Do While startRow <= lastRow
        endRow = startRow + splitCount - 1
        If endRow > lastRow Then endRow = lastRow

        ' 建立新工作表
        Set wsNew = ThisWorkbook.Worksheets.Add
        wsNew.Name = "分組_" & i

        ' 複製標題
        ws.Rows(1).Copy Destination:=wsNew.Rows(1)

        ' 複製資料區塊
        ws.Range(startRow & ":" & endRow).Copy Destination:=wsNew.Rows(2)

        wsNew.Columns.AutoFit
        startRow = endRow + 1
        i = i + 1
    Loop

    MsgBox "拆分完成，共產生 " & (i - 1) & " 個分組工作表。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub
