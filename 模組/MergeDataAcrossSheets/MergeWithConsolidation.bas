Attribute VB_Name = "MergeWithConsolidation"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithConsolidation
'功能說明: 使用 Excel 合併彙算功能，將多個工作表的數值資料
'          依欄位標籤彙整至新的摘要工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub MergeWithConsolidation()
    Dim wb          As Workbook
    Dim summaryWs   As Worksheet
    Dim ws          As Worksheet
    Dim sources()   As String
    Dim srcCount    As Integer
    Dim i           As Integer

    Set wb = ThisWorkbook

    srcCount = 0
    For Each ws In wb.Worksheets
        If ws.Name <> "彙算結果" Then
            srcCount = srcCount + 1
        End If
    Next ws

    If srcCount = 0 Then
        MsgBox "找不到可以合併的工作表。", vbExclamation, "提示"
        Exit Sub
    End If

    ReDim sources(1 To srcCount)
    i = 1
    For Each ws In wb.Worksheets
        If ws.Name <> "彙算結果" Then
            Dim lastRow As Long
            Dim lastCol As Long
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If lastRow >= 2 And lastCol >= 2 Then
                sources(i) = Chr(39) & ws.Name & Chr(39) & "!" & _
                    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Address
                i = i + 1
            End If
        End If
    Next ws

    On Error Resume Next
    Set summaryWs = wb.Worksheets("彙算結果")
    On Error GoTo 0

    If summaryWs Is Nothing Then
        Set summaryWs = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        summaryWs.Name = "彙算結果"
    Else
        summaryWs.Cells.Clear
    End If

    summaryWs.Activate
    summaryWs.Range("A1").Select

    summaryWs.Range("A1").Consolidate _
        Sources:=sources, _
        Function:=xlSum, _
        TopRow:=True, _
        LeftColumn:=True, _
        CreateLinks:=False

    summaryWs.Columns.AutoFit
    MsgBox "合併彙算完成，結果已輸出至「彙算結果」工作表。", vbInformation, "完成"
End Sub

' 建立三個範例工作表用於測試
Sub CreateConsolidationTestData()
    Dim wb      As Workbook
    Dim ws      As Worksheet
    Dim i       As Integer
    Dim names() As String

    Set wb = ThisWorkbook
    names = Split("第一季,第二季,第三季", ",")

    For i = 0 To 2
        On Error Resume Next
        Set ws = wb.Worksheets(names(i))
        On Error GoTo 0

        If ws Is Nothing Then
            Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
            ws.Name = names(i)
        Else
            ws.Cells.Clear
        End If

        ws.Range("A1").Value = "產品"
        ws.Range("B1").Value = "北部"
        ws.Range("C1").Value = "南部"
        ws.Range("A2").Value = "產品A"
        ws.Range("A3").Value = "產品B"
        ws.Range("A4").Value = "產品C"
        ws.Range("B2").Value = (i + 1) * 100
        ws.Range("B3").Value = (i + 1) * 150
        ws.Range("B4").Value = (i + 1) * 80
        ws.Range("C2").Value = (i + 1) * 90
        ws.Range("C3").Value = (i + 1) * 120
        ws.Range("C4").Value = (i + 1) * 110
        ws.Range("A1:C1").Font.Bold = True
        ws.Columns.AutoFit
        Set ws = Nothing
    Next i

    MsgBox "已建立第一季、第二季、第三季三個測試工作表，請執行 MergeWithConsolidation 進行合併彙算。", _
        vbInformation, "測試資料建立完成"
End Sub
