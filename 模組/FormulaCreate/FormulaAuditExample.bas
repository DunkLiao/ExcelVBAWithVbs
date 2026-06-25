Option Explicit
Attribute VB_Name = "FormulaAuditExample"
'*************************************************************************************
'模組名稱: FormulaAuditExample
'功能說明: 示範如何用 VBA 進行公式稽核，顯示前置參照與後置參照
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/25
'
'*************************************************************************************

' 測試用入口
Sub TestFormulaAudit()
    Call AuditFormulaReferences()
End Sub

' 稽核作用儲存格的公式參照
Sub AuditFormulaReferences()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim target As Range
    Dim result As String

    Set ws = GetOrCreateWorksheet("公式稽核範例")
    ws.Cells.Clear

    ' 建立範例資料與公式
    ws.Range("A1").Value = 10
    ws.Range("A2").Value = 20
    ws.Range("A3").Value = 30
    ws.Range("B1").Formula = "=A1*2"
    ws.Range("B2").Formula = "=SUM(A1:A3)"
    ws.Range("B3").Formula = "=AVERAGE(A1:A3)"
    ws.Range("C1").Formula = "=B1+B2"
    ws.Range("C2").Formula = "=B2-B1"
    ws.Range("C3").Formula = "=SUM(B1:B3)"

    ' 檢查每個有公式的儲存格
    Set rng = ws.UsedRange
    result = "公式稽核結果：" & vbCrLf & vbCrLf
    For Each cell In rng
        If cell.HasFormula Then
            result = result & "儲存格 " & cell.Address(False, False) & "：" & vbCrLf
            result = result & "  公式 = " & cell.Formula & vbCrLf

            ' 顯示公式前置參照
            On Error Resume Next
            Set target = cell.Precedents
            If Err.Number = 0 Then
                If target.Count > 0 Then
                    result = result & "  前置參照：" & target.Address(False, False) & vbCrLf
                End If
            End If
            On Error GoTo ErrorHandler

            ' 顯示公式後置參照
            On Error Resume Next
            Set target = cell.Dependents
            If Err.Number = 0 Then
                If target.Count > 0 Then
                    result = result & "  後置參照：" & target.Address(False, False) & vbCrLf
                End If
            End If
            On Error GoTo ErrorHandler
            result = result & vbCrLf
        End If
    Next cell

    ' 顯示稽核結果
    ws.Range("E1").Value = result

    MsgBox "公式稽核完成，結果已寫入工作表。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateWorksheet(ByVal wsName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = wsName
    End If
    Set GetOrCreateWorksheet = ws
End Function
