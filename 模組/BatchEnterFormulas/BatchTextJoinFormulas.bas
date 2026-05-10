Option Explicit
Attribute VB_Name = "BatchTextJoinFormulas"
'*************************************************************************************
'模組名稱: BatchTextJoinFormulas
'功能說明: 批次在指定欄位寫入 TEXTJOIN 公式，合併多欄文字
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Sub BatchTextJoinFormulas()
    Dim ws          As Worksheet
    Dim lastRow     As Long
    Dim startCol    As Integer
    Dim endCol      As Integer
    Dim outputCol   As Integer
    Dim delimiter   As String
    Dim i           As Long
    Dim c           As Integer
    Dim colLetters  As String

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "資料不足，請先確認工作表有資料。", vbExclamation, "提示"
        Exit Sub
    End If

    ' 取得設定
    delimiter = InputBox("請輸入分隔符號（例如 , 或空格）：", "分隔符號", ", ")
    startCol = CInt(InputBox("請輸入起始欄號（數字，例如 1 = A欄）：", "起始欄", "1"))
    endCol = CInt(InputBox("請輸入結束欄號（數字，例如 3 = C欄）：", "結束欄", "3"))
    outputCol = CInt(InputBox("請輸入輸出欄號（數字）：", "輸出欄", CStr(endCol + 1)))

    ' 寫入標題
    ws.Cells(1, outputCol).Value = "TEXTJOIN結果"

    ' 批次寫入公式
    For i = 2 To lastRow
        colLetters = ""
        For c = startCol To endCol
            If colLetters <> "" Then colLetters = colLetters & ", "
            colLetters = colLetters & ws.Cells(i, c).Address(False, False)
        Next c
        ws.Cells(i, outputCol).Formula = _
            "=TEXTJOIN(""" & delimiter & """, TRUE, " & colLetters & ")"
    Next i

    ws.Columns(outputCol).AutoFit
    MsgBox "TEXTJOIN 公式批次寫入完成，共 " & (lastRow - 1) & " 列。", vbInformation, "完成"
End Sub
