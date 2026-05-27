Option Explicit
'*************************************************************************************
'模組名稱: CellRankingFormatting
'功能說明: 依儲存格數值在範圍中的排名，自動套用條件式格式（前N名綠底、後N名紅底）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub ApplyCellRankingFormatting()
    Dim rng As Range
    Dim topN As Long
    Dim bottomN As Long
    Dim userInput As String
    Dim fcTop As FormatCondition
    Dim fcBottom As FormatCondition

    On Error GoTo ErrHandler

    ' 取得使用者選取的數值範圍
    On Error Resume Next
    Set rng = Application.InputBox( _
        "請選取要套用排名格式的數值範圍：", "選取範圍", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

    ' 詢問前幾名
    userInput = InputBox( _
        "請輸入要標示的前幾名（例如輸入 3 代表前3名）：", "設定前幾名", "3")
    If userInput = "" Then Exit Sub
    If Not IsNumeric(userInput) Then
        MsgBox "請輸入有效數字！", vbExclamation, "錯誤"
        Exit Sub
    End If
    topN = CLng(userInput)
    If topN < 1 Then
        MsgBox "前幾名必須大於 0！", vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 詢問後幾名
    userInput = InputBox( _
        "請輸入要標示的後幾名（例如輸入 3 代表後3名）：", "設定後幾名", "3")
    If userInput = "" Then Exit Sub
    If Not IsNumeric(userInput) Then
        MsgBox "請輸入有效數字！", vbExclamation, "錯誤"
        Exit Sub
    End If
    bottomN = CLng(userInput)
    If bottomN < 1 Then
        MsgBox "後幾名必須大於 0！", vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 清除既有條件式格式
    rng.FormatConditions.Delete

    ' 套用「前N名」格式（綠底白字）
    Set fcTop = rng.FormatConditions.AddTop10
    With fcTop
        .TopBottom = xlTop10Top
        .Rank = topN
        .Percent = False
        .Interior.Color = RGB(0, 176, 80)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    ' 套用「後N名」格式（紅底白字）
    Set fcBottom = rng.FormatConditions.AddTop10
    With fcBottom
        .TopBottom = xlTop10Bottom
        .Rank = bottomN
        .Percent = False
        .Interior.Color = RGB(255, 0, 0)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With

    MsgBox "排名格式套用完成！" & vbNewLine & _
           "前 " & topN & " 名：綠底白字" & vbNewLine & _
           "後 " & bottomN & " 名：紅底白字", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

Sub ClearCellRankingFormatting()
    ' 清除選取範圍的排名條件式格式
    Dim rng As Range

    On Error GoTo ErrHandler

    On Error Resume Next
    Set rng = Application.InputBox( _
        "請選取要清除排名格式的範圍：", "選取範圍", Type:=8)
    On Error GoTo ErrHandler
    If rng Is Nothing Then Exit Sub

    rng.FormatConditions.Delete
    MsgBox "已清除選取範圍的所有條件式格式。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
