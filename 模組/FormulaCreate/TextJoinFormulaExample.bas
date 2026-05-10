Option Explicit
Attribute VB_Name = "TextJoinFormulaExample"
'*************************************************************************************
'模組名稱: TextJoinFormulaExample
'功能說明: 以 VBA 批次寫入 TEXTJOIN 公式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Sub CreateTextJoinFormulaExample()
    Dim ws          As Worksheet
    Dim lastRow     As Long
    Dim i           As Long

    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "TextJoinDemo"

    ' 寫入示範資料欄位
    ws.Cells(1, 1).Value = "姓"
    ws.Cells(1, 2).Value = "名"
    ws.Cells(1, 3).Value = "部門"
    ws.Cells(1, 4).Value = "合併全名"

    Dim demo(1 To 5, 1 To 3) As String
    demo(1, 1) = "王": demo(1, 2) = "小明": demo(1, 3) = "業務部"
    demo(2, 1) = "李": demo(2, 2) = "美華": demo(2, 3) = "財務部"
    demo(3, 1) = "張": demo(3, 2) = "建國": demo(3, 3) = "資訊部"
    demo(4, 1) = "陳": demo(4, 2) = "志遠": demo(4, 3) = "人資部"
    demo(5, 1) = "林": demo(5, 2) = "佳蓉": demo(5, 3) = "行政部"

    For i = 1 To 5
        ws.Cells(i + 1, 1).Value = demo(i, 1)
        ws.Cells(i + 1, 2).Value = demo(i, 2)
        ws.Cells(i + 1, 3).Value = demo(i, 3)
    Next i

    lastRow = 6

    ' 批次寫入 TEXTJOIN 公式
    Dim r As Long
    For r = 2 To lastRow
        ws.Cells(r, 4).Formula = _
            "=TEXTJOIN("" "", TRUE, A" & r & ", B" & r & ", ""["" & C" & r & " & ""]"")"
    Next r

    ws.Columns("A:D").AutoFit
    MsgBox "TEXTJOIN 公式已成功寫入！", vbInformation, "完成"
End Sub
