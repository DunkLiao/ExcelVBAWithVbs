Option Explicit
'*************************************************************************************
'模組名稱: TopBottomRankFormatting
'功能說明: 建立前幾名與後幾名的條件格式範例
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/10
'
'*************************************************************************************

Public Sub ApplyTopBottomRankFormatting()
    Dim ws As Worksheet
    Dim scoreRange As Range
    Dim fc As Top10

    On Error GoTo ErrHandler

    Set ws = GetOrCreateTopBottomSheet("排名格式範例")
    ws.Cells.Clear
    Call FillTopBottomData(ws)

    Set scoreRange = ws.Range("C2:C11")
    scoreRange.FormatConditions.Delete

    Set fc = scoreRange.FormatConditions.AddTop10
    With fc
        .TopBottom = xlTop10Top
        .Rank = 3
        .Percent = False
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
        .Font.Bold = True
    End With

    Set fc = scoreRange.FormatConditions.AddTop10
    With fc
        .TopBottom = xlTop10Bottom
        .Rank = 2
        .Percent = False
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
        .Font.Bold = True
    End With

    ws.Columns("A:C").AutoFit
    MsgBox "前 3 名與後 2 名條件格式已建立完成。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "建立排名條件格式失敗：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillTopBottomData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("部門", "員工", "績效分數")
    ws.Range("A2:C2").Value = Array("業務", "王小明", 92)
    ws.Range("A3:C3").Value = Array("業務", "李小華", 68)
    ws.Range("A4:C4").Value = Array("客服", "陳美玲", 81)
    ws.Range("A5:C5").Value = Array("客服", "張志強", 57)
    ws.Range("A6:C6").Value = Array("財務", "林雅婷", 88)
    ws.Range("A7:C7").Value = Array("財務", "周建宏", 74)
    ws.Range("A8:C8").Value = Array("資訊", "吳佩君", 95)
    ws.Range("A9:C9").Value = Array("資訊", "許家豪", 62)
    ws.Range("A10:C10").Value = Array("行政", "黃怡君", 79)
    ws.Range("A11:C11").Value = Array("行政", "鄭柏翰", 49)
End Sub

Private Function GetOrCreateTopBottomSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateTopBottomSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateTopBottomSheet Is Nothing Then
        Set GetOrCreateTopBottomSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateTopBottomSheet.Name = sheetName
    End If
End Function
