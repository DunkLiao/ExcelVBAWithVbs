Attribute VB_Name = "CleanMixedDataTypes"
Option Explicit
'*************************************************************************************
'模組名稱: 清理混合資料類型
'功能說明: 自動偵測並統一修正儲存格中混合的資料類型
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub CleanMixedDataTypes()
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim cell As Range
    Dim fixedCount As Long
    Dim rangeAddr As String
    Dim numVal As Double
    Dim dateVal As Date

    Set ws = ActiveSheet

    rangeAddr = InputBox("請輸入要清理的範圍（例如：A2:C100）：", "選擇範圍", "A2:A100")
    If rangeAddr = "" Then Exit Sub

    On Error Resume Next
    Set targetRange = ws.Range(rangeAddr)
    On Error GoTo 0

    If targetRange Is Nothing Then
        MsgBox "無效的範圍設定。", vbExclamation, "錯誤"
        Exit Sub
    End If

    fixedCount = 0

    For Each cell In targetRange
        If cell.Value <> "" Then
            If cell.HasFormula = False Then
                If IsNumeric(cell.Value) And cell.NumberFormat = "@" Then
                    numVal = CDbl(cell.Value)
                    cell.NumberFormat = "General"
                    cell.Value = numVal
                    fixedCount = fixedCount + 1
                ElseIf IsNumeric(cell.Value) And VarType(cell.Value) = vbString Then
                    cell.Value = CDbl(cell.Value)
                    fixedCount = fixedCount + 1
                ElseIf VarType(cell.Value) = vbString Then
                    If IsDate(cell.Value) Then
                        dateVal = CDate(cell.Value)
                        cell.Value = dateVal
                        cell.NumberFormat = "yyyy/m/d"
                        fixedCount = fixedCount + 1
                    End If
                End If
            End If
        End If
    Next cell

    MsgBox "清理完成！共修正 " & fixedCount & " 個儲存格的資料類型。", vbInformation, "完成"
End Sub

Sub CreateMixedDataDemo()
    Dim ws As Worksheet
    Dim demoName As String

    demoName = "混合資料類型範例"

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(demoName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = demoName
    End If

    ws.Cells.Clear

    ws.Range("A1").Value = "欄位說明"
    ws.Range("B1").Value = "原始資料"

    ws.Range("A2").Value = "文字數字"
    ws.Range("B2").NumberFormat = "@"
    ws.Range("B2").Value = "12345"

    ws.Range("A3").Value = "正常數字"
    ws.Range("B3").Value = 9876

    ws.Range("A4").Value = "文字日期"
    ws.Range("B4").Value = "2026/5/13"

    ws.Range("A5").Value = "正常文字"
    ws.Range("B5").Value = "測試文字"

    ws.Columns("A:B").AutoFit
    ws.Range("A1:B1").Font.Bold = True

    MsgBox "範例資料已建立，請執行 CleanMixedDataTypes 並選取 B2:B5 進行清理。", _
           vbInformation, "提示"
End Sub
