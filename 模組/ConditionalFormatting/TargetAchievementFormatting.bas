Attribute VB_Name = "TargetAchievementFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: TargetAchievementFormatting
'功能說明: 依目標達成率自動套用條件式格式：
'          達成率 < 80%  → 紅底白字（未達標）
'          80% <= 達成率 < 100% → 橙底黑字（接近達標）
'          達成率 >= 100% → 綠底黑字（達標）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

Sub ApplyTargetAchievementFormatting()
    Dim ws          As Worksheet
    Dim rng         As Range
    Dim lastRow     As Long
    Dim rateColIdx  As Long
    Dim rateColAddr As String

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "沒有找到資料列，請確認工作表有資料。", vbExclamation
        Exit Sub
    End If

    ' 輸入達成率欄位
    Dim colInput As String
    colInput = InputBox("請輸入達成率欄位的英文代號（例如 C）：", "設定達成率欄位", "C")
    If colInput = "" Then Exit Sub

    rateColIdx = Range(colInput & "1").Column
    rateColAddr = colInput

    ' 設定格式範圍（第2列到最後一列）
    Set rng = ws.Range(rateColAddr & "2:" & rateColAddr & lastRow)
    rng.FormatConditions.Delete

    ' ── 條件 1：達成率 >= 100% → 綠底黑字
    Dim fc1 As FormatCondition
    Set fc1 = rng.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlGreaterEqual, _
        Formula1:="1")
    fc1.Interior.Color = RGB(0, 176, 80)
    fc1.Font.Color = RGB(0, 0, 0)
    fc1.Font.Bold = True

    ' ── 條件 2：80% <= 達成率 < 100% → 橙底黑字
    Dim fc2 As FormatCondition
    Set fc2 = rng.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlBetween, _
        Formula1:="0.8", _
        Formula2:="0.9999")
    fc2.Interior.Color = RGB(255, 192, 0)
    fc2.Font.Color = RGB(0, 0, 0)

    ' ── 條件 3：達成率 < 80% → 紅底白字
    Dim fc3 As FormatCondition
    Set fc3 = rng.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlLess, _
        Formula1:="0.8")
    fc3.Interior.Color = RGB(192, 0, 0)
    fc3.Font.Color = RGB(255, 255, 255)
    fc3.Font.Bold = True

    ' 確保達成率欄顯示百分比格式
    rng.NumberFormat = "0.0%"

    MsgBox "目標達成率條件式格式已套用！" & vbCrLf & _
           "綠色：達標（>=100%）" & vbCrLf & _
           "橙色：接近（80%~99%）" & vbCrLf & _
           "紅色：未達標（<80%）", vbInformation, "完成"
End Sub

' 建立示範資料並套用格式
Sub CreateTargetAchievementDemo()
    Dim ws As Worksheet
    Set ws = GetOrCreateTargetSheet("達成率格式示範")
    ws.Cells.Clear

    ' 標題
    ws.Range("A1:D1").Value = Array("業務員", "目標金額", "實際銷售", "達成率")

    ' 資料
    ws.Range("A2:D2").Value = Array("張志豪", 100000, 115000, "=C2/B2")
    ws.Range("A3:D3").Value = Array("李佳蓉", 100000, 88000, "=C3/B3")
    ws.Range("A4:D4").Value = Array("王大明", 80000, 75000, "=C4/B4")
    ws.Range("A5:D5").Value = Array("陳雅婷", 90000, 91000, "=C5/B5")
    ws.Range("A6:D6").Value = Array("林志偉", 120000, 95000, "=C6/B6")
    ws.Range("A7:D7").Value = Array("黃淑芬", 70000, 70500, "=C7/B7")

    ' 數值格式
    ws.Range("B2:C7").NumberFormat = "#,##0"
    ws.Range("D2:D7").NumberFormat = "0.0%"

    ' 套用達成率條件格式（D欄）
    Dim rng As Range
    Set rng = ws.Range("D2:D7")
    rng.FormatConditions.Delete

    Dim fc As FormatCondition

    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="1")
    fc.Interior.Color = RGB(0, 176, 80)
    fc.Font.Color = RGB(0, 0, 0)
    fc.Font.Bold = True

    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, _
                                       Formula1:="0.8", Formula2:="0.9999")
    fc.Interior.Color = RGB(255, 192, 0)
    fc.Font.Color = RGB(0, 0, 0)

    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0.8")
    fc.Interior.Color = RGB(192, 0, 0)
    fc.Font.Color = RGB(255, 255, 255)
    fc.Font.Bold = True

    ws.Columns("A:D").AutoFit
    MsgBox "達成率格式示範已建立！", vbInformation, "完成"
End Sub

Private Function GetOrCreateTargetSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateTargetSheet = ws
End Function
