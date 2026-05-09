Attribute VB_Name = "GenerateDiffSummary"
Option Explicit
'*************************************************************************************
'模組名稱: GenerateDiffSummary
'功能說明: 比對兩張工作表並產生差異統計摘要，包含各欄差異數量、
'          差異比率、最大差異值等統計指標
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestGenerateDiffSummary()
    Call CreateSummaryData
    Call GenerateDiffSummary("月報A版", "月報B版", "差異統計摘要")
End Sub

' 建立統計摘要範例資料
Private Sub CreateSummaryData()
    Dim wsA As Worksheet
    Dim wsB As Worksheet

    Set wsA = GetOrCreateSheetGDS("月報A版")
    Set wsB = GetOrCreateSheetGDS("月報B版")

    wsA.Range("A1").Value = "產品"  : wsA.Range("B1").Value = "一月" : wsA.Range("C1").Value = "二月" : wsA.Range("D1").Value = "三月"
    wsA.Range("A2").Value = "產品A" : wsA.Range("B2").Value = 100 : wsA.Range("C2").Value = 200 : wsA.Range("D2").Value = 150
    wsA.Range("A3").Value = "產品B" : wsA.Range("B3").Value = 300 : wsA.Range("C3").Value = 280 : wsA.Range("D3").Value = 310
    wsA.Range("A4").Value = "產品C" : wsA.Range("B4").Value = 50  : wsA.Range("C4").Value = 60  : wsA.Range("D4").Value = 55
    wsA.Range("A5").Value = "產品D" : wsA.Range("B5").Value = 400 : wsA.Range("C5").Value = 420 : wsA.Range("D5").Value = 390
    wsA.Columns("A:D").AutoFit

    wsB.Range("A1").Value = "產品"  : wsB.Range("B1").Value = "一月" : wsB.Range("C1").Value = "二月" : wsB.Range("D1").Value = "三月"
    wsB.Range("A2").Value = "產品A" : wsB.Range("B2").Value = 105 : wsB.Range("C2").Value = 200 : wsB.Range("D2").Value = 160
    wsB.Range("A3").Value = "產品B" : wsB.Range("B3").Value = 300 : wsB.Range("C3").Value = 295 : wsB.Range("D3").Value = 310
    wsB.Range("A4").Value = "產品C" : wsB.Range("B4").Value = 48  : wsB.Range("C4").Value = 60  : wsB.Range("D4").Value = 55
    wsB.Range("A5").Value = "產品D" : wsB.Range("B5").Value = 400 : wsB.Range("C5").Value = 420 : wsB.Range("D5").Value = 400
    wsB.Columns("A:D").AutoFit
End Sub

' 產生差異統計摘要報表
Public Sub GenerateDiffSummary(ByVal sheetA As String, ByVal sheetB As String, _
                                ByVal reportSheet As String)
    Dim wsA          As Worksheet
    Dim wsB          As Worksheet
    Dim wsR          As Worksheet
    Dim lastRow      As Long
    Dim lastCol      As Long
    Dim r            As Long
    Dim c            As Long
    Dim rptRow       As Long
    Dim valA         As Double
    Dim valB         As Double
    Dim colDiffCount As Long
    Dim totalDiff    As Long
    Dim maxDiff      As Double
    Dim curDiff      As Double
    Dim colHeader    As String

    On Error GoTo ErrHandler

    Set wsA = ThisWorkbook.Worksheets(sheetA)
    Set wsB = ThisWorkbook.Worksheets(sheetB)
    Set wsR = GetOrCreateSheetGDS(reportSheet)

    lastRow = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    lastCol = wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column

    ' 報表標頭
    wsR.Range("A1").Value = "差異統計摘要報表"
    With wsR.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
    End With
    wsR.Range("A2").Value = "比對來源: " & sheetA & " vs " & sheetB
    wsR.Range("A3").Value = "產生日期: " & Format(Now, "yyyy/mm/dd hh:mm:ss")

    wsR.Range("A5").Value = "欄位名稱"
    wsR.Range("B5").Value = "差異筆數"
    wsR.Range("C5").Value = "差異比率"
    wsR.Range("D5").Value = "最大差異值"
    wsR.Range("E5").Value = "差異方向"
    With wsR.Range("A5:E5")
        .Font.Bold = True
        .Interior.Color = RGB(47, 117, 181)
        .Font.Color = RGB(255, 255, 255)
    End With

    rptRow = 6
    totalDiff = 0

    For c = 2 To lastCol
        colHeader = CStr(wsA.Cells(1, c).Value)
        colDiffCount = 0
        maxDiff = 0

        For r = 2 To lastRow
            If IsNumeric(wsA.Cells(r, c).Value) And IsNumeric(wsB.Cells(r, c).Value) Then
                valA = CDbl(wsA.Cells(r, c).Value)
                valB = CDbl(wsB.Cells(r, c).Value)
                If valA <> valB Then
                    colDiffCount = colDiffCount + 1
                    curDiff = Abs(valB - valA)
                    If curDiff > maxDiff Then maxDiff = curDiff
                End If
            ElseIf CStr(wsA.Cells(r, c).Value) <> CStr(wsB.Cells(r, c).Value) Then
                colDiffCount = colDiffCount + 1
            End If
        Next r

        wsR.Cells(rptRow, 1).Value = colHeader
        wsR.Cells(rptRow, 2).Value = colDiffCount
        wsR.Cells(rptRow, 3).Value = Format(colDiffCount / (lastRow - 1), "0.0%")
        If maxDiff > 0 Then
            wsR.Cells(rptRow, 4).Value = maxDiff
            wsR.Cells(rptRow, 5).Value = "數值差異"
        Else
            wsR.Cells(rptRow, 4).Value = "-"
            If colDiffCount > 0 Then
                wsR.Cells(rptRow, 5).Value = "文字差異"
            Else
                wsR.Cells(rptRow, 5).Value = "無差異"
            End If
        End If
        If colDiffCount > 0 Then
            wsR.Cells(rptRow, 1).Resize(1, 5).Interior.Color = RGB(255, 235, 156)
        End If
        totalDiff = totalDiff + colDiffCount
        rptRow = rptRow + 1
    Next c

    ' 總計列
    wsR.Cells(rptRow, 1).Value = "合計"
    wsR.Cells(rptRow, 2).Value = totalDiff
    With wsR.Cells(rptRow, 1).Resize(1, 5)
        .Font.Bold = True
        .Interior.Color = RGB(155, 194, 230)
    End With

    wsR.Columns("A:E").AutoFit
    wsR.Activate
    MsgBox "差異統計摘要產生完成！總差異筆數: " & totalDiff, vbInformation, "統計結果"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetGDS(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetGDS = ws
End Function
