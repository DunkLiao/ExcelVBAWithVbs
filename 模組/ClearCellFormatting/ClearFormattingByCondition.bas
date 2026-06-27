Attribute VB_Name = "ClearFormattingByCondition"
Option Explicit
'*************************************************************************************
'模組名稱: ClearFormattingByCondition
'功能說明: 根據儲存格值條件選擇性清除格式的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestClearFormattingByCondition()
    Call ClearFormattingIfConditionMet
End Sub

Sub ClearFormattingIfConditionMet()
    Dim ws As Worksheet
    Dim cell As Range
    Dim clearCount As Integer
    Dim safeThreshold As Double

    On Error Resume Next
    Application.DisplayAlerts = False
    Set ws = ThisWorkbook.Worksheets("選擇性清除格式")
    If Not ws Is Nothing Then ws.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "選擇性清除格式"

    ' 填入含格式的資料
    ws.Range("A1").Value = "項目"
    ws.Range("B1").Value = "狀態"
    ws.Range("C1").Value = "餘額"
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(68, 114, 196)
    ws.Range("A1:C1").Font.Color = RGB(255, 255, 255)

    ws.Range("A2").Value = "帳戶A"
    ws.Range("B2").Value = "正常"
    ws.Range("C2").Value = 50000

    ws.Range("A3").Value = "帳戶B"
    ws.Range("B3").Value = "逾期"
    ws.Range("C3").Value = -2500
    ws.Range("A3:C3").Interior.Color = RGB(255, 199, 206)

    ws.Range("A4").Value = "帳戶C"
    ws.Range("B4").Value = "正常"
    ws.Range("C4").Value = 32000

    ws.Range("A5").Value = "帳戶D"
    ws.Range("B5").Value = "逾期"
    ws.Range("C5").Value = -800
    ws.Range("A5:C5").Interior.Color = RGB(255, 199, 206)

    ws.Columns("A:C").AutoFit

    ' 設定安全門檻：只清除逾期且餘額低於門檻的儲存格格式
    safeThreshold = 0

    clearCount = 0
    Dim r As Long
    For r = 2 To 5
        Set cell = ws.Cells(r, 3)

        If ws.Cells(r, 2).Value = "逾期" And cell.Value < safeThreshold Then
            ' 只清除背景色和字型格式，保留數值
            cell.Interior.ColorIndex = xlNone
            cell.Font.ColorIndex = xlAutomatic
            cell.Font.Bold = False

            ws.Cells(r, 1).Interior.ColorIndex = xlNone
            ws.Cells(r, 1).Font.ColorIndex = xlAutomatic
            ws.Cells(r, 1).Font.Bold = False

            ws.Cells(r, 2).Interior.ColorIndex = xlNone
            ws.Cells(r, 2).Font.ColorIndex = xlAutomatic
            ws.Cells(r, 2).Font.Bold = False

            clearCount = clearCount + 1
        End If
    Next r

    MsgBox "選擇性清除格式完成！" & vbCrLf & _
           "已清除 " & clearCount & " 個逾期項目的格式。" & vbCrLf & _
           "正常項目格式保持不變。", vbInformation, "完成"
End Sub
