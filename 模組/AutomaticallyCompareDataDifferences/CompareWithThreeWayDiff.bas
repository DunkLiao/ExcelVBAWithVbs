Attribute VB_Name = "CompareWithThreeWayDiff"
Option Explicit
'*************************************************************************************
'模組名稱: CompareWithThreeWayDiff
'功能說明: 三方差異比對（基準版 vs 版本A vs 版本B），產出差異摘要的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestCompareWithThreeWayDiff()
    Call ThreeWayDataComparison
End Sub

Sub ThreeWayDataComparison()
    Dim wsBase As Worksheet
    Dim wsA As Worksheet
    Dim wsB As Worksheet
    Dim wsResult As Worksheet

    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsBase = ThisWorkbook.Worksheets("基準版")
    If Not wsBase Is Nothing Then wsBase.Delete
    Set wsA = ThisWorkbook.Worksheets("版本A")
    If Not wsA Is Nothing Then wsA.Delete
    Set wsB = ThisWorkbook.Worksheets("版本B")
    If Not wsB Is Nothing Then wsB.Delete
    Set wsResult = ThisWorkbook.Worksheets("三方比對結果")
    If Not wsResult Is Nothing Then wsResult.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 建立基準版資料
    Set wsBase = ThisWorkbook.Worksheets.Add
    wsBase.Name = "基準版"
    wsBase.Range("A1").Value = "項目"
    wsBase.Range("B1").Value = "預算"
    wsBase.Range("A2").Value = "行銷"
    wsBase.Range("B2").Value = 100000
    wsBase.Range("A3").Value = "人事"
    wsBase.Range("B3").Value = 200000
    wsBase.Range("A4").Value = "研發"
    wsBase.Range("B4").Value = 150000
    wsBase.Range("A1:B1").Font.Bold = True

    ' 建立版本A
    Set wsA = ThisWorkbook.Worksheets.Add
    wsA.Name = "版本A"
    wsA.Range("A1").Value = "項目"
    wsA.Range("B1").Value = "預算"
    wsA.Range("A2").Value = "行銷"
    wsA.Range("B2").Value = 120000
    wsA.Range("A3").Value = "人事"
    wsA.Range("B3").Value = 200000
    wsA.Range("A4").Value = "研發"
    wsA.Range("B4").Value = 140000
    wsA.Range("A1:B1").Font.Bold = True

    ' 建立版本B
    Set wsB = ThisWorkbook.Worksheets.Add
    wsB.Name = "版本B"
    wsB.Range("A1").Value = "項目"
    wsB.Range("B1").Value = "預算"
    wsB.Range("A2").Value = "行銷"
    wsB.Range("B2").Value = 110000
    wsB.Range("A3").Value = "人事"
    wsB.Range("B3").Value = 210000
    wsB.Range("A4").Value = "研發"
    wsB.Range("B4").Value = 150000
    wsB.Range("A1:B1").Font.Bold = True

    ' 建立結果工作表
    Set wsResult = ThisWorkbook.Worksheets.Add
    wsResult.Name = "三方比對結果"

    wsResult.Range("A1").Value = "項目"
    wsResult.Range("B1").Value = "基準版"
    wsResult.Range("C1").Value = "版本A"
    wsResult.Range("D1").Value = "版本B"
    wsResult.Range("E1").Value = "A差異"
    wsResult.Range("F1").Value = "B差異"
    wsResult.Range("G1").Value = "差異方向"
    wsResult.Range("A1:G1").Font.Bold = True

    Dim i As Integer
    Dim baseVal As Double
    Dim valA As Double
    Dim valB As Double
    Dim diffA As Double
    Dim diffB As Double
    Dim direction As String

    For i = 2 To 4
        wsResult.Cells(i, 1).Value = wsBase.Cells(i, 1).Value
        wsResult.Cells(i, 2).Value = wsBase.Cells(i, 2).Value
        wsResult.Cells(i, 3).Value = wsA.Cells(i, 2).Value
        wsResult.Cells(i, 4).Value = wsB.Cells(i, 2).Value

        baseVal = wsBase.Cells(i, 2).Value
        valA = wsA.Cells(i, 2).Value
        valB = wsB.Cells(i, 2).Value

        diffA = valA - baseVal
        diffB = valB - baseVal

        wsResult.Cells(i, 5).Value = diffA
        wsResult.Cells(i, 6).Value = diffB

        ' 判斷差異方向
        If diffA > 0 And diffB > 0 Then
            direction = "兩版皆增"
        ElseIf diffA < 0 And diffB < 0 Then
            direction = "兩版皆減"
        ElseIf diffA = 0 And diffB = 0 Then
            direction = "無差異"
        ElseIf diffA > 0 And diffB < 0 Then
            direction = "A增B減"
        ElseIf diffA < 0 And diffB > 0 Then
            direction = "A減B增"
        ElseIf diffA <> 0 And diffB = 0 Then
            direction = "僅A有差異"
        ElseIf diffA = 0 And diffB <> 0 Then
            direction = "僅B有差異"
        End If
        wsResult.Cells(i, 7).Value = direction
    Next i

    wsResult.Columns("A:G").AutoFit

    MsgBox "三方差異比對完成！請查看「三方比對結果」工作表。", vbInformation, "完成"
End Sub
