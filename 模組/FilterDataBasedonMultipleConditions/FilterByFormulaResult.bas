'*************************************************************************************
'模組名稱: FilterByFormulaResult
'功能說明: 依據公式計算結果進行篩選，將符合條件的列複製至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************
Option Explicit

Sub FilterByFormulaResult()
    Dim ws          As Worksheet
    Dim wsResult    As Worksheet
    Dim lastRow     As Long
    Dim lastCol     As Long
    Dim i           As Long
    Dim resultRow   As Long
    Dim checkCol    As Long
    Dim threshold   As Double
    Dim colName     As String
    Dim inputVal    As String

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "資料不足，請確認工作表有資料。", vbExclamation, "提示"
        Exit Sub
    End If

    ' 取得要判斷的欄位名稱
    colName = InputBox("請輸入要篩選的欄位名稱（標題）：", "篩選設定")
    If colName = "" Then Exit Sub

    inputVal = InputBox("請輸入篩選門檻值（大於此值的列才保留）：", "篩選設定")
    If inputVal = "" Then Exit Sub
    If Not IsNumeric(inputVal) Then
        MsgBox "門檻值必須為數字！", vbExclamation, "錯誤"
        Exit Sub
    End If
    threshold = CDbl(inputVal)

    ' 找欄位索引
    checkCol = 0
    Dim c As Long
    For c = 1 To lastCol
        If ws.Cells(1, c).Value = colName Then
            checkCol = c
            Exit For
        End If
    Next c

    If checkCol = 0 Then
        MsgBox "找不到欄位：" & colName, vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 建立結果工作表
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("篩選結果").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsResult = ThisWorkbook.Sheets.Add( _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResult.Name = "篩選結果"

    ' 複製標題
    ws.Rows(1).Copy wsResult.Rows(1)
    resultRow = 2

    ' 逐列以數值門檻判斷
    For i = 2 To lastRow
        Dim cellVal As Double
        On Error Resume Next
        cellVal = CDbl(ws.Cells(i, checkCol).Value)
        On Error GoTo 0
        If cellVal > threshold Then
            ws.Rows(i).Copy wsResult.Rows(resultRow)
            resultRow = resultRow + 1
        End If
    Next i

    wsResult.Columns.AutoFit
    MsgBox "篩選完成！欄位 " & colName & " > " & threshold & vbCrLf & _
           "共保留 " & (resultRow - 2) & " 筆資料，輸出至篩選結果工作表。", _
           vbInformation, "完成"
End Sub
