Attribute VB_Name = "MergeWithAutoNaming"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithAutoNaming
'功能說明: 跨表合併資料時，依日期自動命名結果工作表（如Merge_20260531）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestMergeWithAutoNaming()
    Call CreateSampleSourceSheets
    Call MergeWithAutoNaming
End Sub

Sub MergeWithAutoNaming()
    Dim wb         As Workbook
    Dim srcWs      As Worksheet
    Dim destWs     As Worksheet
    Dim destName   As String
    Dim destRow    As Long
    Dim lastRow    As Long
    Dim lastCol    As Long
    Dim hasHeader  As Boolean

    Set wb = ThisWorkbook
    destName = "Merge_" & Format(Now, "YYYYMMDD")

    On Error Resume Next
    Set destWs = wb.Worksheets(destName)
    On Error GoTo 0
    If Not destWs Is Nothing Then
        destName = destName & "_" & Format(Now, "HHMMSS")
        Set destWs = Nothing
    End If

    Set destWs = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    destWs.Name = destName

    destRow = 1
    hasHeader = False

    Application.ScreenUpdating = False

    For Each srcWs In wb.Worksheets
        If srcWs.Name <> destName And Left(srcWs.Name, 3) = "src" Then
            lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
            lastCol = srcWs.UsedRange.Columns.Count

            If lastRow >= 1 Then
                If Not hasHeader Then
                    srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(lastRow, lastCol)).Copy _
                        Destination:=destWs.Cells(destRow, 1)
                    destRow = destRow + lastRow
                    hasHeader = True
                Else
                    If lastRow > 1 Then
                        srcWs.Range(srcWs.Cells(2, 1), srcWs.Cells(lastRow, lastCol)).Copy _
                            Destination:=destWs.Cells(destRow, 1)
                        destRow = destRow + lastRow - 1
                    End If
                End If
            End If
        End If
    Next srcWs

    Application.ScreenUpdating = True
    destWs.Columns.AutoFit

    MsgBox "自動命名合併完成！結果表：" & destName & Chr(13) & _
           "共合併 " & destRow - 1 & " 列資料。", vbInformation, "完成"
End Sub

Private Sub CreateSampleSourceSheets()
    Dim wb As Workbook
    Dim ws As Worksheet
    Set wb = ThisWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets("src1")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = "src1"
    End If
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("姓名", "部門", "業績")
    ws.Range("A2:C2").Value = Array("王大明", "業務部", 120000)
    ws.Range("A3:C3").Value = Array("李小華", "業務部", 98000)

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("src2")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = "src2"
    End If
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("姓名", "部門", "業績")
    ws.Range("A2:C2").Value = Array("陳美玲", "行銷部", 85000)
    ws.Range("A3:C3").Value = Array("張志偉", "行銷部", 110000)
End Sub
