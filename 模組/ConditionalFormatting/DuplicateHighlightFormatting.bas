Attribute VB_Name = "DuplicateHighlightFormatting"
Option Explicit

' ============================================================
' 模組名稱：DuplicateHighlightFormatting
' 功能說明：對選取範圍套用條件式格式，醒目提示重複值
'           提供三種模式：全欄範圍 / 選取範圍 / 多欄重複
' ============================================================

Sub HighlightDuplicateValues()
    Dim rng         As Range
    Dim choice      As String
    Dim ws          As Worksheet
    
    On Error GoTo ErrHandler
    
    ' 詢問使用者模式
    choice = InputBox("請選擇醒目提示模式：" & vbCrLf & _
                      "1 = 選取範圍中的重複值" & vbCrLf & _
                      "2 = 清除選取範圍的重複值格式" & vbCrLf & _
                      "3 = 跨欄比較重複（A 欄 vs B 欄）", _
                      "重複值醒目提示", "1")
    
    If choice = "" Then
        MsgBox "已取消操作。", vbInformation, "取消"
        Exit Sub
    End If
    
    Select Case choice
        Case "1"
            Call ApplyDuplicateFormat
        Case "2"
            Call ClearDuplicateFormat
        Case "3"
            Call HighlightCrossColumnDuplicates
        Case Else
            MsgBox "請輸入 1、2 或 3。", vbExclamation, "輸入錯誤"
    End Select
    
    Exit Sub
    
ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub

' 套用重複值條件式格式
Private Sub ApplyDuplicateFormat()
    Dim rng As Range
    Set rng = Application.InputBox( _
        "請選取要檢查重複值的範圍：", "選取範圍", Type:=8)
    
    If rng Is Nothing Then Exit Sub
    
    ' 清除該範圍原有的條件式格式
    rng.FormatConditions.Delete
    
    ' 新增重複值條件式格式
    Dim fc As FormatCondition
    Set fc = rng.FormatConditions.AddUniqueValues()
    fc.DupeUnique = xlDuplicate
    
    ' 設定醒目顯示樣式（橘底深紅字）
    With fc.Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 199, 206)
        .TintAndShade = 0
    End With
    With fc.Font
        .Color = RGB(156, 0, 6)
        .Bold = True
    End With
    
    MsgBox "重複值格式已套用！共 " & rng.Cells.Count & " 個儲存格受影響。", _
           vbInformation, "完成"
End Sub

' 清除重複值條件式格式
Private Sub ClearDuplicateFormat()
    Dim rng As Range
    Set rng = Application.InputBox( _
        "請選取要清除重複值格式的範圍：", "選取範圍", Type:=8)
    
    If rng Is Nothing Then Exit Sub
    
    rng.FormatConditions.Delete
    MsgBox "條件式格式已清除。", vbInformation, "完成"
End Sub

' 跨欄比較重複：醒目提示 A 欄中也出現在 B 欄的值
Private Sub HighlightCrossColumnDuplicates()
    Dim ws          As Worksheet
    Dim colA        As Range
    Dim colB        As Range
    Dim cellA       As Range
    Dim lastRowA    As Long
    Dim lastRowB    As Long
    
    Set ws = ActiveSheet
    lastRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastRowB = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    
    If lastRowA < 1 Or lastRowB < 1 Then
        MsgBox "A 欄或 B 欄沒有資料。", vbExclamation, "提示"
        Exit Sub
    End If
    
    Set colA = ws.Range("A2:A" & lastRowA)
    Set colB = ws.Range("B2:B" & lastRowB)
    
    ' 清除 A 欄現有條件式格式
    colA.FormatConditions.Delete
    
    ' 使用公式型條件式格式
    Dim fc As Object
    Set fc = colA.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=COUNTIF($B$2:$B$" & lastRowB & ",A2)>0")
    
    With fc.Interior
        .Color = RGB(255, 199, 206)
    End With
    With fc.Font
        .Color = RGB(156, 0, 6)
        .Bold = True
    End With
    
    MsgBox "跨欄重複比較格式套用完成！" & vbCrLf & _
           "A 欄中同時出現在 B 欄的值已標示紅色。", vbInformation, "完成"
End Sub