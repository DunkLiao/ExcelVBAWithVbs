Attribute VB_Name = "ClearSparklineFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearSparklineFormatting
'功能說明: 清除現用工作表中所有走勢圖（Sparkline）群組，並可選擇是否保留欄位資料
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestClearSparklineFormatting()
    Call ClearAllSparklines
End Sub

' 清除工作表中所有走勢圖
Sub ClearAllSparklines()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim clearedCount As Long
    Dim answer As Integer

    Set ws = ActiveSheet
    clearedCount = 0

    If ws.SparklineGroups.Count = 0 Then
        MsgBox "此工作表中沒有任何走勢圖。", vbInformation, "提示"
        Exit Sub
    End If

    answer = MsgBox("找到 " & ws.SparklineGroups.Count & " 個走勢圖群組。" & vbCrLf & _
                    "確定要清除所有走勢圖嗎？", vbYesNo + vbQuestion, "確認清除")
    If answer <> vbYes Then
        MsgBox "已取消，走勢圖未清除。", vbInformation, "取消"
        Exit Sub
    End If

    Do While ws.SparklineGroups.Count > 0
        ws.SparklineGroups(1).Delete
        clearedCount = clearedCount + 1
    Loop

    MsgBox "已清除 " & clearedCount & " 個走勢圖群組！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "清除走勢圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 建立示範走勢圖並測試清除
Sub CreateAndClearSparklineExample()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(ThisWorkbook, "走勢圖示範")

    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "一月"
    ws.Range("C1").Value = "二月"
    ws.Range("D1").Value = "三月"
    ws.Range("E1").Value = "四月"
    ws.Range("F1").Value = "走勢圖"
    ws.Range("A1:F1").Font.Bold = True

    ws.Range("A2").Value = "北部"
    ws.Range("B2").Value = 120
    ws.Range("C2").Value = 145
    ws.Range("D2").Value = 98
    ws.Range("E2").Value = 167

    ws.Range("A3").Value = "中部"
    ws.Range("B3").Value = 88
    ws.Range("C3").Value = 102
    ws.Range("D3").Value = 135
    ws.Range("E3").Value = 119

    ws.Range("A4").Value = "南部"
    ws.Range("B4").Value = 76
    ws.Range("C4").Value = 93
    ws.Range("D4").Value = 110
    ws.Range("E4").Value = 88

    ' 建立折線走勢圖
    ws.Range("F2:F4").SparklineGroups.Add _
        Type:=xlSparkLine, _
        SourceData:=ws.Range("B2:E4").Address

    ws.Columns("A:F").AutoFit

    MsgBox "走勢圖已建立。請執行 ClearAllSparklines 以清除走勢圖。", vbInformation, "提示"
    Exit Sub

ErrorHandler:
    MsgBox "建立走勢圖時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
