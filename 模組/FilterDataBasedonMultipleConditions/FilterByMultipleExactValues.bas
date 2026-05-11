Attribute VB_Name = "FilterByMultipleExactValues"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByMultipleExactValues
'功能說明: 依據使用者輸入的多個精確值，篩選指定欄位並將結果輸出至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

Sub FilterByMultipleExactValues()
    Dim ws As Worksheet
    Dim resultWs As Worksheet
    Dim filterCol As Integer
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long
    Dim destRow As Long
    Dim inputStr As String
    Dim values() As String
    Dim v As Integer
    Dim cellVal As String
    Dim matched As Boolean
    Dim resultName As String

    Set ws = ActiveSheet

    filterCol = CInt(InputBox("請輸入篩選欄位的欄號（例如：1代表A欄）：", "設定篩選欄", "1"))
    If filterCol < 1 Then Exit Sub

    inputStr = InputBox("請輸入要篩選的精確值，多個值以逗號分隔：" & vbCrLf & _
        "例如：台北,台中,高雄", "輸入篩選值")
    If inputStr = "" Then Exit Sub

    values = Split(inputStr, ",")
    For v = 0 To UBound(values)
        values(v) = Trim(values(v))
    Next v

    lastRow = ws.Cells(ws.Rows.Count, filterCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "工作表資料不足。", vbExclamation, "提示"
        Exit Sub
    End If

    resultName = "精確值篩選結果"

    On Error Resume Next
    Set resultWs = ThisWorkbook.Worksheets(resultName)
    On Error GoTo 0

    If resultWs Is Nothing Then
        Set resultWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        resultWs.Name = resultName
    Else
        resultWs.Cells.Clear
    End If

    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy resultWs.Range("A1")
    destRow = 2

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        cellVal = Trim(CStr(ws.Cells(i, filterCol).Value))
        matched = False
        For v = 0 To UBound(values)
            If cellVal = values(v) Then
                matched = True
                Exit For
            End If
        Next v
        If matched Then
            ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Copy resultWs.Cells(destRow, 1)
            destRow = destRow + 1
        End If
    Next i

    Application.ScreenUpdating = True
    resultWs.Columns.AutoFit

    MsgBox "篩選完成！符合條件共 " & (destRow - 2) & " 筆。", vbInformation, "完成"
End Sub
