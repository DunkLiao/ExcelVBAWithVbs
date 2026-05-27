Option Explicit
'*************************************************************************************
'模組名稱: BatchSubstituteFormulas
'功能說明: 批次在目標範圍輸入 SUBSTITUTE 函數，將來源範圍中的指定字元批次取代
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub BatchSubstituteFormulas()
    Dim srcRng As Range
    Dim destRng As Range
    Dim srcCell As Range
    Dim destCell As Range
    Dim oldText As String
    Dim newText As String
    Dim instanceNum As String
    Dim formulaStr As String
    Dim i As Long

    On Error GoTo ErrHandler

    ' 取得來源範圍
    On Error Resume Next
    Set srcRng = Application.InputBox( _
        "請選取包含原始文字的來源範圍：", "選取來源範圍", Type:=8)
    On Error GoTo ErrHandler
    If srcRng Is Nothing Then Exit Sub

    ' 輸入要取代的舊字元
    oldText = InputBox("請輸入要被取代的文字（舊字元）：", "設定舊字元", " ")

    ' 輸入取代後的新字元
    newText = InputBox("請輸入取代後的文字（新字元，可留空）：", "設定新字元", "_")

    ' 輸入取代第幾個出現（留空表示全部）
    instanceNum = InputBox( _
        "請輸入要取代第幾個出現（留空表示全部取代）：", "設定取代次數", "")

    ' 取得目標範圍（放公式的位置）
    On Error Resume Next
    Set destRng = Application.InputBox( _
        "請選取貼入 SUBSTITUTE 公式的目標範圍（左上角即可）：", "選取目標範圍", Type:=8)
    On Error GoTo ErrHandler
    If destRng Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    i = 0

    For Each srcCell In srcRng
        ' 計算目標儲存格位置
        Set destCell = destRng.Cells(1).Offset( _
            Int(i / srcRng.Columns.Count), _
            i Mod srcRng.Columns.Count)

        ' 組合 SUBSTITUTE 公式
        If Trim(instanceNum) = "" Then
            ' 全部取代
            formulaStr = "=SUBSTITUTE(" & srcCell.Address & "," & _
                Chr(34) & oldText & Chr(34) & "," & _
                Chr(34) & newText & Chr(34) & ")"
        Else
            ' 只取代指定第幾次出現
            formulaStr = "=SUBSTITUTE(" & srcCell.Address & "," & _
                Chr(34) & oldText & Chr(34) & "," & _
                Chr(34) & newText & Chr(34) & "," & _
                Trim(instanceNum) & ")"
        End If

        destCell.Formula = formulaStr
        i = i + 1
    Next srcCell

    Application.ScreenUpdating = True

    MsgBox "SUBSTITUTE 公式批次輸入完成！共處理 " & i & " 個儲存格。", _
           vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

Sub BatchSubstituteFormulasDemo()
    ' 在作用工作表建立 SUBSTITUTE 示範資料與公式
    Dim ws As Worksheet
    Dim i As Long

    On Error GoTo ErrHandler

    Set ws = ActiveSheet

    ws.Cells(1, 1).Value = "原始文字"
    ws.Cells(1, 2).Value = "取代空格為底線"
    ws.Cells(1, 3).Value = "移除逗號"
    ws.Cells(1, 4).Value = "取代第1個空格"
    ws.Rows(1).Font.Bold = True

    Dim demos(1 To 5) As String
    demos(1) = "Hello World Excel"
    demos(2) = "台灣 Excel VBA 教學"
    demos(3) = "apple,banana,cherry"
    demos(4) = "2026 05 27"
    demos(5) = "A B C D E"

    For i = 1 To 5
        ws.Cells(i + 1, 1).Value = demos(i)
        ws.Cells(i + 1, 2).Formula = _
            "=SUBSTITUTE(A" & (i + 1) & "," & Chr(34) & " " & Chr(34) & "," & Chr(34) & "_" & Chr(34) & ")"
        ws.Cells(i + 1, 3).Formula = _
            "=SUBSTITUTE(A" & (i + 1) & "," & Chr(34) & "," & Chr(34) & "," & Chr(34) & Chr(34) & ")"
        ws.Cells(i + 1, 4).Formula = _
            "=SUBSTITUTE(A" & (i + 1) & "," & Chr(34) & " " & Chr(34) & "," & Chr(34) & "_" & Chr(34) & ",1)"
    Next i

    ws.Columns("A:D").AutoFit
    MsgBox "SUBSTITUTE 示範資料建立完成！", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
