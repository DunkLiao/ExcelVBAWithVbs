Attribute VB_Name = "FilterByNumericRange"
Option Explicit

' ============================================================
' 模組名稱：FilterByNumericRange
' 功能說明：依使用者輸入的數值上下限，篩選指定欄位的資料
'           支援：大於/小於/介於 三種模式
'           篩選結果可輸出至新工作表
' ============================================================

Sub FilterByNumericRange()
    Dim ws          As Worksheet
    Dim wsResult    As Worksheet
    Dim lastRow     As Long
    Dim lastCol     As Long
    Dim filterCol   As Long
    Dim minVal      As Double
    Dim maxVal      As Double
    Dim mode        As String
    Dim resultName  As String
    Dim nextRow     As Long
    Dim i           As Long
    Dim cellVal     As Double
    Dim matched     As Boolean
    Dim colInput    As String
    Dim minInput    As String
    Dim maxInput    As String
    
    On Error GoTo ErrHandler
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    If lastRow < 2 Then
        MsgBox "工作表中沒有足夠的資料。", vbExclamation, "提示"
        Exit Sub
    End If
    
    ' 輸入篩選欄號
    colInput = InputBox("請輸入要篩選的欄號（數字，如 2 代表 B 欄）：", "篩選欄號", "2")
    If colInput = "" Then Exit Sub
    If Not IsNumeric(colInput) Then
        MsgBox "請輸入有效的欄號數字。", vbExclamation, "輸入錯誤"
        Exit Sub
    End If
    filterCol = CLng(colInput)
    If filterCol < 1 Or filterCol > lastCol Then
        MsgBox "欄號超出資料範圍（1 到 " & lastCol & "）。", vbExclamation, "輸入錯誤"
        Exit Sub
    End If
    
    ' 輸入篩選模式
    mode = InputBox("請選擇篩選模式：" & vbCrLf & _
                    "1 = 介於下限和上限之間（含端點）" & vbCrLf & _
                    "2 = 大於等於下限" & vbCrLf & _
                    "3 = 小於等於上限", _
                    "篩選模式", "1")
    If mode = "" Then Exit Sub
    If mode <> "1" And mode <> "2" And mode <> "3" Then
        MsgBox "請輸入 1、2 或 3。", vbExclamation, "輸入錯誤"
        Exit Sub
    End If
    
    ' 依模式輸入數值
    If mode = "1" Or mode = "2" Then
        minInput = InputBox("請輸入下限值（最小值）：", "下限", "0")
        If minInput = "" Then Exit Sub
        If Not IsNumeric(minInput) Then
            MsgBox "請輸入有效數字。", vbExclamation, "輸入錯誤"
            Exit Sub
        End If
        minVal = CDbl(minInput)
    End If
    
    If mode = "1" Or mode = "3" Then
        maxInput = InputBox("請輸入上限值（最大值）：", "上限", "100")
        If maxInput = "" Then Exit Sub
        If Not IsNumeric(maxInput) Then
            MsgBox "請輸入有效數字。", vbExclamation, "輸入錯誤"
            Exit Sub
        End If
        maxVal = CDbl(maxInput)
    End If
    
    If mode = "1" And minVal > maxVal Then
        MsgBox "下限值不能大於上限值。", vbExclamation, "輸入錯誤"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 建立篩選結果工作表
    resultName = "數值範圍篩選結果"
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(resultName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True
    
    Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResult.Name = resultName
    
    ' 複製標題列
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy _
        Destination:=wsResult.Range("A1")
    wsResult.Rows(1).Font.Bold = True
    wsResult.Rows(1).Interior.Color = RGB(68, 114, 196)
    wsResult.Rows(1).Font.Color = RGB(255, 255, 255)
    
    ' 在標題列後加上篩選條件說明
    Dim condDesc As String
    Select Case mode
        Case "1"
            condDesc = ws.Cells(1, filterCol).Value & " 介於 " & minVal & " 和 " & maxVal & " 之間"
        Case "2"
            condDesc = ws.Cells(1, filterCol).Value & " >= " & minVal
        Case "3"
            condDesc = ws.Cells(1, filterCol).Value & " <= " & maxVal
    End Select
    wsResult.Cells(1, lastCol + 1).Value = "篩選條件：" & condDesc
    wsResult.Cells(1, lastCol + 1).Font.Italic = True
    wsResult.Cells(1, lastCol + 1).Font.Color = RGB(255, 255, 255)
    
    nextRow = 2
    
    ' 逐列篩選
    For i = 2 To lastRow
        Dim rawVal As Variant
        rawVal = ws.Cells(i, filterCol).Value
        
        If IsNumeric(rawVal) Then
            cellVal = CDbl(rawVal)
            matched = False
            
            Select Case mode
                Case "1"
                    matched = (cellVal >= minVal And cellVal <= maxVal)
                Case "2"
                    matched = (cellVal >= minVal)
                Case "3"
                    matched = (cellVal <= maxVal)
            End Select
            
            If matched Then
                ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Copy _
                    Destination:=wsResult.Cells(nextRow, 1)
                nextRow = nextRow + 1
            End If
        End If
    Next i
    
    wsResult.Columns.AutoFit
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Dim matchCount As Long
    matchCount = nextRow - 2
    MsgBox "數值範圍篩選完成！" & vbCrLf & _
           "篩選條件：" & condDesc & vbCrLf & _
           "共篩選出 " & matchCount & " 筆資料。" & vbCrLf & _
           "結果已存至「" & resultName & "」工作表。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub