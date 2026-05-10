'*************************************************************************************
'模組名稱: SplitByKeyword
'功能說明: 依據指定欄位中的關鍵字，將資料列切割至不同工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************
Option Explicit

Sub SplitByKeyword()
    Dim ws          As Worksheet
    Dim wsNew       As Worksheet
    Dim lastRow     As Long
    Dim lastCol     As Long
    Dim i           As Long
    Dim keyword     As String
    Dim colIndex    As Long
    Dim colName     As String
    Dim dict        As Object

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' 詢問要依哪欄關鍵字切割
    colName = InputBox("請輸入要作為切割依據的欄位名稱（第一列標題）：", "切割依據")
    If colName = "" Then Exit Sub

    ' 找欄位索引
    colIndex = 0
    Dim c As Long
    For c = 1 To lastCol
        If ws.Cells(1, c).Value = colName Then
            colIndex = c
            Exit For
        End If
    Next c

    If colIndex = 0 Then
        MsgBox "找不到欄位：" & colName, vbExclamation, "錯誤"
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    ' 收集所有關鍵字
    For i = 2 To lastRow
        keyword = CStr(ws.Cells(i, colIndex).Value)
        If Not dict.Exists(keyword) Then
            dict.Add keyword, keyword
        End If
    Next i

    ' 為每個關鍵字建立工作表並複製標題
    Dim key As Variant
    For Each key In dict.Keys
        Dim shName As String
        shName = Left(CStr(key), 31)
        On Error Resume Next
        Set wsNew = ThisWorkbook.Sheets(shName)
        On Error GoTo 0
        If wsNew Is Nothing Then
            Set wsNew = ThisWorkbook.Sheets.Add( _
                After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsNew.Name = shName
            ws.Rows(1).Copy wsNew.Rows(1)
        End If
        Set wsNew = Nothing
    Next key

    ' 複製資料列到對應工作表
    For i = 2 To lastRow
        keyword = CStr(ws.Cells(i, colIndex).Value)
        shName = Left(keyword, 31)
        Set wsNew = ThisWorkbook.Sheets(shName)
        Dim tgtRow As Long
        tgtRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row + 1
        ws.Rows(i).Copy wsNew.Rows(tgtRow)
    Next i

    MsgBox "依關鍵字切割完成，共分出 " & dict.Count & " 個工作表。", vbInformation, "完成"
End Sub
