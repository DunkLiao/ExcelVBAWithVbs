Attribute VB_Name = "MergeWithFormulaLinks"
Option Explicit

'*************************************************************************************
'模組名稱: MergeWithFormulaLinks
'功能說明: 跨工作表以公式參照方式彙整資料
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub MergeWithFormulaLinks()
    Dim wsMerge As Worksheet
    Dim ws As Worksheet
    Dim destRow As Long
    Dim srcLastRow As Long
    Dim i As Long
    Dim sheetName As String

    On Error Resume Next
    Set wsMerge = ThisWorkbook.Worksheets("彙整")
    On Error GoTo 0

    If wsMerge Is Nothing Then
        Set wsMerge = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsMerge.Name = "彙整"
    Else
        wsMerge.Cells.Clear
    End If

    wsMerge.Range("A1").Value = "來源工作表"
    wsMerge.Range("B1").Value = "列號"
    wsMerge.Range("C1").Value = "欄A資料（公式參照）"
    wsMerge.Range("A1:C1").Font.Bold = True
    destRow = 2

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "彙整" Then
            sheetName = ws.Name
            srcLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            For i = 1 To srcLastRow
                wsMerge.Cells(destRow, 1).Value = sheetName
                wsMerge.Cells(destRow, 2).Value = i
                wsMerge.Cells(destRow, 3).Formula = _
                    "='" & sheetName & "'!A" & i
                destRow = destRow + 1
            Next i
        End If
    Next ws

    wsMerge.Columns("A:C").AutoFit
    MsgBox "彙整完成，共 " & (destRow - 2) & " 筆資料以公式參照！", vbInformation
End Sub
