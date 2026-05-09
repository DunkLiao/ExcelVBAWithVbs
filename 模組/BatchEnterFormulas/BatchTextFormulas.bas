Attribute VB_Name = "BatchTextFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchTextFormulas
'功能說明: 批次填入文字處理公式，包含 LEFT、RIGHT、MID、LEN、TRIM、UPPER、LOWER
'
'作者版權: Dunk
'原始設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestBatchTextFormulas()
    Call CreateTextFormulaExample
End Sub

' 建立文字公式批次填入示範
Sub CreateTextFormulaExample()
    Dim ws As Worksheet
    On Error GoTo ErrHandler

    Set ws = GetOrCreateTextSheet(ThisWorkbook, "文字公式示範")
    Call FillRawTextData(ws)
    Call BatchEnterTextExtractFormulas(ws)
    Call BatchEnterTextTransformFormulas(ws)

    ws.Columns("A:I").AutoFit
    ws.Activate
    MsgBox "文字處理公式已批次填入完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 批次填入 LEFT / RIGHT / MID / LEN 擷取公式
Private Sub BatchEnterTextExtractFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("B1").Value = "前3字(LEFT)"
    ws.Range("C1").Value = "後3字(RIGHT)"
    ws.Range("D1").Value = "中間4字(MID)"
    ws.Range("E1").Value = "字元長度(LEN)"
    ws.Range("B1:E1").Font.Bold = True

    For i = 2 To lastRow
        ws.Cells(i, 2).Formula = "=LEFT(A" & i & ",3)"
        ws.Cells(i, 3).Formula = "=RIGHT(A" & i & ",3)"
        ws.Cells(i, 4).Formula = "=MID(A" & i & ",2,4)"
        ws.Cells(i, 5).Formula = "=LEN(A" & i & ")"
    Next i
End Sub

' 批次填入 TRIM / UPPER / LOWER 轉換公式
Private Sub BatchEnterTextTransformFormulas(ByVal ws As Worksheet)
    Dim i As Integer
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ws.Range("F1").Value = "去除空白(TRIM)"
    ws.Range("G1").Value = "大寫(UPPER)"
    ws.Range("H1").Value = "小寫(LOWER)"
    ws.Range("I1").Value = "合併姓名"
    ws.Range("F1:I1").Font.Bold = True

    For i = 2 To lastRow
        ws.Cells(i, 6).Formula = "=TRIM(A" & i & ")"
        ws.Cells(i, 7).Formula = "=UPPER(A" & i & ")"
        ws.Cells(i, 8).Formula = "=LOWER(A" & i & ")"
        ' CONCATENATE 模擬：將前3字與後3字合併
        ws.Cells(i, 9).Formula = "=LEFT(A" & i & ",3)&""-""&RIGHT(A" & i & ",3)"
    Next i
End Sub

' 填入原始文字資料（英文員工代號）
Private Sub FillRawTextData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "員工代號"
    ws.Range("A1").Font.Bold = True

    ws.Range("A2").Value = "EMP-Alice-2023"
    ws.Range("A3").Value = "EMP-Bob-2022"
    ws.Range("A4").Value = "EMP-Carol-2024"
    ws.Range("A5").Value = "EMP-David-2021"
    ws.Range("A6").Value = "EMP-Eve-2023"
    ws.Range("A7").Value = "EMP-Frank-2020"
    ws.Range("A8").Value = "EMP-Grace-2022"
    ws.Range("A9").Value = "EMP-Henry-2024"
    ws.Range("A10").Value = "EMP-Ivy-2023"
    ws.Range("A11").Value = "EMP-Jack-2021"
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateTextSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateTextSheet = ws
End Function