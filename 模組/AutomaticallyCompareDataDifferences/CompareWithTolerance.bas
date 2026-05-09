Attribute VB_Name = "CompareWithTolerance"
Option Explicit
'*************************************************************************************
'模組名稱: CompareWithTolerance
'功能說明: 數值欄位容差比對，允許設定可接受誤差百分比，
'          超出容差範圍才標記為差異
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口，容差設定為 5%
Sub TestCompareWithTolerance()
    Call CreateToleranceData
    Call CompareWithTolerance("預算表", "實績表", "容差比對報告", 0.05)
End Sub

' 建立容差比對範例資料
Private Sub CreateToleranceData()
    Dim wsA As Worksheet
    Dim wsB As Worksheet

    Set wsA = GetOrCreateSheetCWT("預算表")
    Set wsB = GetOrCreateSheetCWT("實績表")

    wsA.Range("A1").Value = "部門"
    wsA.Range("B1").Value = "Q1預算"
    wsA.Range("C1").Value = "Q2預算"
    wsA.Range("D1").Value = "Q3預算"
    wsA.Range("A2").Value = "業務部" : wsA.Range("B2").Value = 100000 : wsA.Range("C2").Value = 120000 : wsA.Range("D2").Value = 110000
    wsA.Range("A3").Value = "研發部" : wsA.Range("B3").Value = 200000 : wsA.Range("C3").Value = 210000 : wsA.Range("D3").Value = 205000
    wsA.Range("A4").Value = "行銷部" : wsA.Range("B4").Value = 80000  : wsA.Range("C4").Value = 90000  : wsA.Range("D4").Value = 85000
    wsA.Columns("A:D").AutoFit

    wsB.Range("A1").Value = "部門"
    wsB.Range("B1").Value = "Q1預算"
    wsB.Range("C1").Value = "Q2預算"
    wsB.Range("D1").Value = "Q3預算"
    wsB.Range("A2").Value = "業務部" : wsB.Range("B2").Value = 103000 : wsB.Range("C2").Value = 125000 : wsB.Range("D2").Value = 110000
    wsB.Range("A3").Value = "研發部" : wsB.Range("B3").Value = 218000 : wsB.Range("C3").Value = 209000 : wsB.Range("D3").Value = 205000
    wsB.Range("A4").Value = "行銷部" : wsB.Range("B4").Value = 80000  : wsB.Range("C4").Value = 100000 : wsB.Range("D4").Value = 85500
    wsB.Columns("A:D").AutoFit
End Sub

' 進行容差比對，tolerancePct 為允許誤差比率（如 0.05 代表 5%）
Public Sub CompareWithTolerance(ByVal sheetA As String, ByVal sheetB As String, _
                                 ByVal reportSheet As String, ByVal tolerancePct As Double)
    Dim wsA      As Worksheet
    Dim wsB      As Worksheet
    Dim wsR      As Worksheet
    Dim lastRow  As Long
    Dim lastCol  As Long
    Dim r        As Long
    Dim c        As Long
    Dim rptRow   As Long
    Dim valA     As Double
    Dim valB     As Double
    Dim diffPct  As Double
    Dim diffCount As Long
    Dim isNumA   As Boolean
    Dim isNumB   As Boolean

    On Error GoTo ErrHandler

    Set wsA = ThisWorkbook.Worksheets(sheetA)
    Set wsB = ThisWorkbook.Worksheets(sheetB)
    Set wsR = GetOrCreateSheetCWT(reportSheet)

    lastRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    lastCol = wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column

    wsR.Range("A1").Value = "儲存格位址"
    wsR.Range("B1").Value = "欄位/列說明"
    wsR.Range("C1").Value = sheetA & " 值"
    wsR.Range("D1").Value = sheetB & " 值"
    wsR.Range("E1").Value = "差異百分比"
    wsR.Range("F1").Value = "是否超出容差"
    With wsR.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(255, 165, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    rptRow = 2
    diffCount = 0

    For r = 2 To lastRow
        For c = 2 To lastCol
            isNumA = IsNumeric(wsA.Cells(r, c).Value)
            isNumB = IsNumeric(wsB.Cells(r, c).Value)
            If isNumA And isNumB Then
                valA = CDbl(wsA.Cells(r, c).Value)
                valB = CDbl(wsB.Cells(r, c).Value)
                If valA <> 0 Then
                    diffPct = Abs((valB - valA) / valA)
                ElseIf valB <> 0 Then
                    diffPct = 1
                Else
                    diffPct = 0
                End If
                wsR.Cells(rptRow, 1).Value = wsA.Cells(r, c).Address(False, False)
                wsR.Cells(rptRow, 2).Value = CStr(wsA.Cells(1, c).Value) & " / " & CStr(wsA.Cells(r, 1).Value)
                wsR.Cells(rptRow, 3).Value = valA
                wsR.Cells(rptRow, 4).Value = valB
                wsR.Cells(rptRow, 5).Value = Format(diffPct, "0.00%")
                If diffPct > tolerancePct Then
                    wsR.Cells(rptRow, 6).Value = "超出"
                    wsR.Cells(rptRow, 1).Resize(1, 6).Interior.Color = RGB(255, 199, 206)
                    diffCount = diffCount + 1
                Else
                    wsR.Cells(rptRow, 6).Value = "正常"
                End If
                rptRow = rptRow + 1
            End If
        Next c
    Next r

    wsR.Columns("A:F").AutoFit
    wsR.Activate
    MsgBox "容差比對完成！" & vbCrLf & _
           "容差設定: " & Format(tolerancePct, "0%") & vbCrLf & _
           "超出容差: " & diffCount & " 筆", vbInformation, "比對結果"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetCWT(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetCWT = ws
End Function
