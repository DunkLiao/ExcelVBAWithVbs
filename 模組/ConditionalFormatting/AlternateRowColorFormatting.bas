Attribute VB_Name = "AlternateRowColorFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: AlternateRowColorFormatting
'功能說明: 以條件式格式化公式，為資料範圍套用交替列背景色的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點
Sub TestAlternateRowColorFormatting()
    Dim ws As Worksheet
    Set ws = GetOrCreateAltRowSheet("交替列色彩範例")
    Call FillAltRowSampleData(ws)
    Call ApplyAlternateRowColor(ws, ws.Range("A1:E10"), _
                                RGB(189, 215, 238), RGB(255, 255, 255))
    MsgBox "交替列色彩條件格式已套用完成！", vbInformation, "完成"
End Sub

' 對指定範圍套用交替列色彩條件格式
' ws: 目標工作表
' targetRange: 要套用格式的範圍
' evenColor: 偶數列背景色
' oddColor: 奇數列背景色
Sub ApplyAlternateRowColor(ByVal ws As Worksheet, ByVal targetRange As Range, _
                            ByVal evenColor As Long, ByVal oddColor As Long)
    On Error GoTo ErrorHandler

    targetRange.FormatConditions.Delete

    Dim fcOdd As FormatCondition
    Set fcOdd = targetRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=MOD(ROW(),2)=1")
    fcOdd.Interior.Color = oddColor

    Dim fcEven As FormatCondition
    Set fcEven = targetRange.FormatConditions.Add( _
        Type:=xlExpression, _
        Formula1:="=MOD(ROW(),2)=0")
    fcEven.Interior.Color = evenColor

    Exit Sub

ErrorHandler:
    MsgBox "套用交替列色彩時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 清除指定範圍的交替列條件格式
Sub ClearAlternateRowColor(ByVal targetRange As Range)
    targetRange.FormatConditions.Delete
    MsgBox "交替列色彩條件格式已清除。", vbInformation, "完成"
End Sub

' 填入範例資料
Private Sub FillAltRowSampleData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:E1").Value = Array("編號", "姓名", "部門", "職稱", "薪資")
    ws.Range("A1:E1").Font.Bold = True

    Dim names  As Variant
    Dim depts  As Variant
    Dim titles As Variant
    names  = Array("陳大明", "林小華", "王建國", "張美玲", "李志偉", "林淡芬", "黃俊傑", "劉雅婷", "蔡宗翰")
    depts  = Array("業務部", "行銀部", "技術部", "行政部", "業務部", "技術部", "行銀部", "行政部", "技術部")
    titles = Array("專員", "主任", "工程師", "助理", "組長", "高級工程師", "設計師", "文書", "架構師")

    Dim i As Integer
    For i = 1 To 9
        ws.Cells(i + 1, 1).Value = i
        ws.Cells(i + 1, 2).Value = names(i - 1)
        ws.Cells(i + 1, 3).Value = depts(i - 1)
        ws.Cells(i + 1, 4).Value = titles(i - 1)
        ws.Cells(i + 1, 5).Value = 30000 + i * 3000
    Next i

    ws.Columns("A:E").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateAltRowSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateAltRowSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateAltRowSheet Is Nothing Then
        Set GetOrCreateAltRowSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateAltRowSheet.Name = sheetName
    End If
End Function
