Attribute VB_Name = "CompareWithCommentAnnotation"
Option Explicit
'*************************************************************************************
'模組名稱: CompareWithCommentAnnotation
'功能說明: 比較兩個工作表的相同儲存格，在差異處自動新增批註說明新舊值
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口
Sub TestCompareWithCommentAnnotation()
    If ThisWorkbook.Worksheets.Count < 2 Then
        MsgBox "請確保活頁簿中至少有 2 個工作表。", vbExclamation, "錯誤"
        Exit Sub
    End If
    Call CompareWithCommentAnnotation( _
        ThisWorkbook.Worksheets(1), _
        ThisWorkbook.Worksheets(2))
End Sub

' 比較兩個工作表並在差異儲存格新增批註
' ws1: 舊版工作表（基準）
' ws2: 新版工作表（比較目標，差異批註寫入此表）
Sub CompareWithCommentAnnotation( _
    ByVal ws1 As Worksheet, _
    ByVal ws2 As Worksheet)

    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Integer
    Dim oldVal As String
    Dim newVal As String
    Dim diffCount As Long
    Dim cmt As Comment

    lastRow = Application.WorksheetFunction.Max( _
        ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row, _
        ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row)

    lastCol = Application.WorksheetFunction.Max( _
        ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column, _
        ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column)

    If lastRow < 1 Or lastCol < 1 Then
        MsgBox "工作表中無有效資料。", vbExclamation, "錯誤"
        Exit Sub
    End If

    diffCount = 0

    For r = 1 To lastRow
        For c = 1 To lastCol
            oldVal = CStr(ws1.Cells(r, c).Value)
            newVal = CStr(ws2.Cells(r, c).Value)

            If oldVal <> newVal Then
                diffCount = diffCount + 1

                On Error Resume Next
                Set cmt = ws2.Cells(r, c).Comment
                On Error GoTo 0

                If cmt Is Nothing Then
                    Set cmt = ws2.Cells(r, c).AddComment
                End If

                cmt.Visible = False
                cmt.Text Text:="[差異]" & vbCrLf & _
                              "舊值：" & oldVal & vbCrLf & _
                              "新值：" & newVal

                ws2.Cells(r, c).Interior.Color = RGB(255, 255, 153)
            End If
        Next c
    Next r

    MsgBox "比較完成，共發現 " & diffCount & " 處差異。" & vbCrLf & _
           "差異儲存格已標記黃色並加入批註。", vbInformation, "比較結果"
End Sub