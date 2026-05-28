Attribute VB_Name = "CellOutlineFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: CellOutlineFormatting
'功能說明: 依條件對儲存格套用外框線格式，以不同框線顏色區分數值範圍
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Sub TestCellOutlineFormatting()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("外框線條件格式")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "外框線條件格式"
    End If
    ws.Cells.Clear
    Call FillOutlineData(ws)
    Call ApplyCellOutlineFormatting(ws.Range("B2:D11"))
    ws.Columns("A:E").AutoFit
    MsgBox "外框線條件格式已套用完畢！", vbInformation, "完成"
End Sub

' 依數值範圍套用外框線格式
' >= 80 => 綠色粗框；40~79 => 橘色框；< 40 => 紅色框
Sub ApplyCellOutlineFormatting(ByVal targetRange As Range)
    Dim cell      As Range
    Dim cellVal   As Double
    Dim bdrColor  As Long
    Dim bdrWeight As XlBorderWeight

    targetRange.Borders.LineStyle = xlNone

    For Each cell In targetRange
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            cellVal = CDbl(cell.Value)
            If cellVal >= 80 Then
                bdrColor  = RGB(0, 180, 0)
                bdrWeight = xlThick
            ElseIf cellVal >= 40 Then
                bdrColor  = RGB(255, 140, 0)
                bdrWeight = xlMedium
            Else
                bdrColor  = RGB(220, 0, 0)
                bdrWeight = xlThin
            End If
            With cell.Borders(xlEdgeTop)
                .LineStyle = xlContinuous
                .Color = bdrColor
                .Weight = bdrWeight
            End With
            With cell.Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = bdrColor
                .Weight = bdrWeight
            End With
            With cell.Borders(xlEdgeLeft)
                .LineStyle = xlContinuous
                .Color = bdrColor
                .Weight = bdrWeight
            End With
            With cell.Borders(xlEdgeRight)
                .LineStyle = xlContinuous
                .Color = bdrColor
                .Weight = bdrWeight
            End With
        End If
    Next cell
End Sub

Private Sub FillOutlineData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("學生", "國文", "數學", "英文")
    With ws.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Range("A2:D2").Value = Array("王小明", 92, 35, 78)
    ws.Range("A3:D3").Value = Array("李大華", 55, 88, 42)
    ws.Range("A4:D4").Value = Array("張美玲", 71, 65, 95)
    ws.Range("A5:D5").Value = Array("陳志偉", 28, 91, 66)
    ws.Range("A6:D6").Value = Array("林欣怡", 85, 47, 33)
    ws.Range("A7:D7").Value = Array("黃建國", 63, 72, 84)
    ws.Range("A8:D8").Value = Array("吳雅雯", 44, 56, 77)
    ws.Range("A9:D9").Value = Array("鄭宏達", 98, 81, 90)
    ws.Range("A10:D10").Value = Array("許淑芬", 37, 29, 58)
    ws.Range("A11:D11").Value = Array("謝明輝", 76, 93, 61)
End Sub
