Attribute VB_Name = "CleanEmoticonData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanEmoticonData
'功能說明: 清除儲存格文字中的 ASCII 顏文字符號（如 :) :D ^^ :-( 等），
'          保留一般中英文與數字
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Sub TestCleanEmoticonData()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("表情符號清理")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "表情符號清理"
    End If
    ws.Cells.Clear
    Call FillEmoticonData(ws)
    Call CleanEmoticonFromRange(ws.Range("B2:B9"))
    ws.Columns("A:C").AutoFit
    MsgBox "表情符號清理完畢！", vbInformation, "完成"
End Sub

' 清除範圍內的 ASCII 顏文字（如 :) :D ^_^ 等常見組合）
Sub CleanEmoticonFromRange(ByVal rng As Range)
    Dim cell    As Range
    Dim cleaned As String

    Application.ScreenUpdating = False
    For Each cell In rng
        If VarType(cell.Value) = vbString And Len(cell.Value) > 0 Then
            cleaned = RemoveAsciiEmoticons(cell.Value)
            If cleaned <> cell.Value Then
                cell.Value = Trim(cleaned)
            End If
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

' 移除常見 ASCII 顏文字組合
Private Function RemoveAsciiEmoticons(ByVal inputStr As String) As String
    Dim result As String
    result = inputStr

    ' 常見顏文字清單（由長到短排列，避免短式先被替換導致殘留）
    Dim emoticons As Variant
    emoticons = Array( _
        ":-)", ":-(", ":-D", ":-/", ":-|", ":-P", ":-O", _
        ":)", ":(", ":D", ":/", ":|", ":P", ":O", _
        "^_^", "^-^", "(^_^)", "(*^_^*)", _
        "(^^)", "(-_-)", "(>_<)", _
        "XD", "xD", "T_T", "T-T", _
        "=D", "=)", "=(", _
        "^^", "~~")

    Dim i As Integer
    For i = 0 To UBound(emoticons)
        result = Replace(result, emoticons(i), "")
    Next i

    RemoveAsciiEmoticons = result
End Function

Private Sub FillEmoticonData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("訂單編號", "客戶留言", "狀態")
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("A2:C2").Value = Array("ORD-001", "商品很棒！非常滿意 :) ", "已出貨")
    ws.Range("A3:C3").Value = Array("ORD-002", "配送速度超快 (^^) 感謝！", "已簽收")
    ws.Range("A4:C4").Value = Array("ORD-003", "包裝完整 :-D 下次還來！", "已完成")
    ws.Range("A5:C5").Value = Array("ORD-004", "有小瑕疵 :-/ 希望改善", "處理中")
    ws.Range("A6:C6").Value = Array("ORD-005", "價格實惠 :) 品質不錯", "已完成")
    ws.Range("A7:C7").Value = Array("ORD-006", "客服態度良好 =D 讚！", "已完成")
    ws.Range("A8:C8").Value = Array("ORD-007", "等待時間有點長 :-(", "已出貨")
    ws.Range("A9:C9").Value = Array("ORD-008", "整體還不錯 ^^ 推薦！", "已完成")
    ws.Columns("A:C").AutoFit
End Sub
