Attribute VB_Name = "CompareIgnoreCase"
Option Explicit
'*************************************************************************************
'模組名稱: CompareIgnoreCase
'功能說明: 不區分英文大小寫比對兩張工作表，
'          區分「僅大小寫差異」與「實質內容差異」兩類
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestCompareIgnoreCase()
    Call CreateIgnoreCaseData
    Call CompareIgnoreCase("系統匯出", "手動輸入", "大小寫差異報告")
End Sub

' 建立大小寫比對範例資料
Private Sub CreateIgnoreCaseData()
    Dim wsA As Worksheet
    Dim wsB As Worksheet

    Set wsA = GetOrCreateSheetCIC("系統匯出")
    Set wsB = GetOrCreateSheetCIC("手動輸入")

    wsA.Range("A1").Value = "料號"
    wsA.Range("B1").Value = "型號"
    wsA.Range("C1").Value = "顏色代碼"
    wsA.Range("D1").Value = "規格"
    wsA.Range("A2").Value = "SKU001" : wsA.Range("B2").Value = "MODEL-A100" : wsA.Range("C2").Value = "RED"   : wsA.Range("D2").Value = "XL"
    wsA.Range("A3").Value = "SKU002" : wsA.Range("B3").Value = "MODEL-B200" : wsA.Range("C3").Value = "BLUE"  : wsA.Range("D3").Value = "M"
    wsA.Range("A4").Value = "SKU003" : wsA.Range("B4").Value = "MODEL-C300" : wsA.Range("C4").Value = "GREEN" : wsA.Range("D4").Value = "S"
    wsA.Columns("A:D").AutoFit

    wsB.Range("A1").Value = "料號"
    wsB.Range("B1").Value = "型號"
    wsB.Range("C1").Value = "顏色代碼"
    wsB.Range("D1").Value = "規格"
    wsB.Range("A2").Value = "SKU001" : wsB.Range("B2").Value = "model-a100" : wsB.Range("C2").Value = "Red"   : wsB.Range("D2").Value = "XL"
    wsB.Range("A3").Value = "SKU002" : wsB.Range("B3").Value = "MODEL-B200" : wsB.Range("C3").Value = "BLUE"  : wsB.Range("D3").Value = "m"
    wsB.Range("A4").Value = "SKU003" : wsB.Range("B4").Value = "Model-C999" : wsB.Range("C4").Value = "GREEN" : wsB.Range("D4").Value = "S"
    wsB.Columns("A:D").AutoFit
End Sub

' 不區分大小寫比對，區分「大小寫差異」與「實質差異」
Public Sub CompareIgnoreCase(ByVal sheetA As String, ByVal sheetB As String, _
                              ByVal reportSheet As String)
    Dim wsA           As Worksheet
    Dim wsB           As Worksheet
    Dim wsR           As Worksheet
    Dim lastRow       As Long
    Dim lastCol       As Long
    Dim r             As Long
    Dim c             As Long
    Dim rptRow        As Long
    Dim valA          As String
    Dim valB          As String
    Dim caseOnlyCount As Long
    Dim realDiffCount As Long

    On Error GoTo ErrHandler

    Set wsA = ThisWorkbook.Worksheets(sheetA)
    Set wsB = ThisWorkbook.Worksheets(sheetB)
    Set wsR = GetOrCreateSheetCIC(reportSheet)

    lastRow = Application.WorksheetFunction.Max( _
        wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row, _
        wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row)
    lastCol = Application.WorksheetFunction.Max( _
        wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column, _
        wsB.Cells(1, wsB.Columns.Count).End(xlToLeft).Column)

    wsR.Range("A1").Value = "儲存格"
    wsR.Range("B1").Value = "差異類型"
    wsR.Range("C1").Value = sheetA & " 值"
    wsR.Range("D1").Value = sheetB & " 值"
    With wsR.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(192, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    rptRow = 2
    caseOnlyCount = 0
    realDiffCount = 0

    For r = 1 To lastRow
        For c = 1 To lastCol
            valA = CStr(wsA.Cells(r, c).Value)
            valB = CStr(wsB.Cells(r, c).Value)
            If valA <> valB Then
                wsR.Cells(rptRow, 1).Value = wsA.Cells(r, c).Address(False, False)
                wsR.Cells(rptRow, 3).Value = valA
                wsR.Cells(rptRow, 4).Value = valB
                If UCase(valA) = UCase(valB) Then
                    ' 僅大小寫不同
                    wsR.Cells(rptRow, 2).Value = "大小寫差異"
                    wsR.Cells(rptRow, 1).Resize(1, 4).Interior.Color = RGB(255, 235, 156)
                    caseOnlyCount = caseOnlyCount + 1
                Else
                    ' 實質內容不同
                    wsR.Cells(rptRow, 2).Value = "實質差異"
                    wsR.Cells(rptRow, 1).Resize(1, 4).Interior.Color = RGB(255, 199, 206)
                    realDiffCount = realDiffCount + 1
                End If
                rptRow = rptRow + 1
            End If
        Next c
    Next r

    wsR.Columns("A:D").AutoFit
    wsR.Activate
    MsgBox "不區分大小寫比對完成！" & vbCrLf & _
           "僅大小寫差異: " & caseOnlyCount & " 筆" & vbCrLf & _
           "實質內容差異: " & realDiffCount & " 筆", vbInformation, "比對結果"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetCIC(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetCIC = ws
End Function
