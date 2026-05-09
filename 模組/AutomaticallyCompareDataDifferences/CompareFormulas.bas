Attribute VB_Name = "CompareFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: CompareFormulas
'功能說明: 比對兩張工作表中的公式內容（非儲存格值），
'          找出公式不一致的位置並輸出報告
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestCompareFormulas()
    Call CreateFormulaData
    Call CompareFormulas("公式版A", "公式版B", "公式差異報告")
End Sub

' 建立含有公式的範例工作表
Private Sub CreateFormulaData()
    Dim wsA As Worksheet
    Dim wsB As Worksheet

    Set wsA = GetOrCreateSheetCF("公式版A")
    Set wsB = GetOrCreateSheetCF("公式版B")

    wsA.Range("A1").Value = "月份"
    wsA.Range("B1").Value = "收入"
    wsA.Range("C1").Value = "支出"
    wsA.Range("D1").Value = "淨利"
    wsA.Range("A2").Value = "一月" : wsA.Range("B2").Value = 100000 : wsA.Range("C2").Value = 60000
    wsA.Range("A3").Value = "二月" : wsA.Range("B3").Value = 120000 : wsA.Range("C3").Value = 70000
    wsA.Range("D2").Formula = "=B2-C2"
    wsA.Range("D3").Formula = "=B3-C3"
    wsA.Columns("A:D").AutoFit

    wsB.Range("A1").Value = "月份"
    wsB.Range("B1").Value = "收入"
    wsB.Range("C1").Value = "支出"
    wsB.Range("D1").Value = "淨利"
    wsB.Range("A2").Value = "一月" : wsB.Range("B2").Value = 100000 : wsB.Range("C2").Value = 60000
    wsB.Range("A3").Value = "二月" : wsB.Range("B3").Value = 120000 : wsB.Range("C3").Value = 70000
    ' 故意放入不同公式以產生差異
    wsB.Range("D2").Formula = "=B2-C2"
    wsB.Range("D3").Formula = "=B3+C3"
    wsB.Columns("A:D").AutoFit
End Sub

' 比對兩張工作表中的公式，差異輸出至報告工作表
Public Sub CompareFormulas(ByVal sheetA As String, ByVal sheetB As String, _
                            ByVal reportSheet As String)
    Dim wsA      As Worksheet
    Dim wsB      As Worksheet
    Dim wsR      As Worksheet
    Dim lastRow  As Long
    Dim lastCol  As Long
    Dim r        As Long
    Dim c        As Long
    Dim rptRow   As Long
    Dim fmlA     As String
    Dim fmlB     As String
    Dim diffCount As Long

    On Error GoTo ErrHandler

    Set wsA = ThisWorkbook.Worksheets(sheetA)
    Set wsB = ThisWorkbook.Worksheets(sheetB)
    Set wsR = GetOrCreateSheetCF(reportSheet)

    lastRow = Application.WorksheetFunction.Max( _
        wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row, _
        wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row)
    lastCol = Application.WorksheetFunction.Max( _
        wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column, _
        wsB.Cells(1, wsB.Columns.Count).End(xlToLeft).Column)

    wsR.Range("A1").Value = "儲存格"
    wsR.Range("B1").Value = "差異類型"
    wsR.Range("C1").Value = sheetA & " 公式"
    wsR.Range("D1").Value = sheetB & " 公式"
    With wsR.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(112, 48, 160)
        .Font.Color = RGB(255, 255, 255)
    End With

    rptRow = 2
    diffCount = 0

    For r = 1 To lastRow
        For c = 1 To lastCol
            If wsA.Cells(r, c).HasFormula Then
                fmlA = wsA.Cells(r, c).Formula
            Else
                fmlA = ""
            End If
            If wsB.Cells(r, c).HasFormula Then
                fmlB = wsB.Cells(r, c).Formula
            Else
                fmlB = ""
            End If
            If fmlA <> fmlB Then
                wsR.Cells(rptRow, 1).Value = wsA.Cells(r, c).Address(False, False)
                If fmlA <> "" And fmlB = "" Then
                    wsR.Cells(rptRow, 2).Value = "B缺少公式"
                ElseIf fmlA = "" And fmlB <> "" Then
                    wsR.Cells(rptRow, 2).Value = "A缺少公式"
                Else
                    wsR.Cells(rptRow, 2).Value = "公式不同"
                End If
                wsR.Cells(rptRow, 3).Value = fmlA
                wsR.Cells(rptRow, 4).Value = fmlB
                wsR.Cells(rptRow, 1).Resize(1, 4).Interior.Color = RGB(255, 235, 156)
                rptRow = rptRow + 1
                diffCount = diffCount + 1
            End If
        Next c
    Next r

    wsR.Columns("A:D").AutoFit
    wsR.Activate
    If diffCount = 0 Then
        MsgBox "兩張工作表公式完全相符！", vbInformation, "公式比對"
    Else
        MsgBox "公式比對完成！共發現 " & diffCount & " 處公式差異。", vbInformation, "公式比對"
    End If
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetCF(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetCF = ws
End Function
