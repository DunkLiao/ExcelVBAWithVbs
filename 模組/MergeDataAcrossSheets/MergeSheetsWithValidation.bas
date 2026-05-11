Attribute VB_Name = "MergeSheetsWithValidation"
Option Explicit
'*************************************************************************************
'模組名稱: MergeSheetsWithValidation
'功能說明: 合併活頁簿中所有工作表資料，並驗證必填欄位不得為空，匯整至摘要工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口：驗證第 1 欄不得為空並合併
Sub TestMergeSheetsWithValidation()
    Call MergeSheetsWithValidation("合併驗證結果", 1)
End Sub

' 合併所有工作表資料，並驗證指定欄位不得為空
' destSheetName : 目標工作表名稱
' requiredCol   : 必填欄位欄號 (1=A)
Sub MergeSheetsWithValidation( _
    ByVal destSheetName As String, _
    ByVal requiredCol As Integer)

    Dim destWs As Worksheet
    Dim srcWs As Worksheet
    Dim destRow As Long
    Dim srcLastRow As Long
    Dim srcLastCol As Long
    Dim r As Long
    Dim c As Integer
    Dim isFirstSheet As Boolean
    Dim errorCount As Long
    Dim validCount As Long
    Dim errorLog As String
    Dim startRow As Long
    Dim cellVal As String

    On Error Resume Next
    Set destWs = ThisWorkbook.Worksheets(destSheetName)
    On Error GoTo 0

    If destWs Is Nothing Then
        Set destWs = ThisWorkbook.Worksheets.Add
        destWs.Name = destSheetName
    Else
        destWs.Cells.Clear
    End If

    destRow = 1
    isFirstSheet = True
    errorCount = 0
    validCount = 0
    errorLog = ""

    For Each srcWs In ThisWorkbook.Worksheets
        If srcWs.Name = destSheetName Then GoTo NextSheet

        srcLastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
        srcLastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column
        If srcLastRow < 1 Then GoTo NextSheet

        If isFirstSheet Then
            startRow = 1
            isFirstSheet = False
        Else
            startRow = 2
        End If

        For r = startRow To srcLastRow
            If r > 1 Then
                cellVal = Trim(CStr(srcWs.Cells(r, requiredCol).Value))
                If cellVal = "" Then
                    errorCount = errorCount + 1
                    errorLog = errorLog & "工作表[" & srcWs.Name & "] 第 " & r & " 列：必填欄位為空" & vbCrLf
                    GoTo NextRow
                End If
                validCount = validCount + 1
            End If

            For c = 1 To srcLastCol
                destWs.Cells(destRow, c).Value = srcWs.Cells(r, c).Value
            Next c
            destRow = destRow + 1
NextRow:
        Next r
NextSheet:
    Next srcWs

    destWs.Columns.AutoFit

    Dim summary As String
    summary = "合併完成！" & vbCrLf & _
              "有效資料列：" & validCount & " 列" & vbCrLf & _
              "驗證失敗列：" & errorCount & " 列"

    If errorCount > 0 Then
        summary = summary & vbCrLf & vbCrLf & "驗證失敗明細：" & vbCrLf & errorLog
    End If

    MsgBox summary, IIf(errorCount > 0, vbExclamation, vbInformation), "合併驗證結果"
End Sub