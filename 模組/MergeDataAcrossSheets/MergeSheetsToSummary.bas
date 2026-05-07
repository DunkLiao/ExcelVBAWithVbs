Attribute VB_Name = "MergeSheetsToSummary"
Option Explicit
'*************************************************************************************
'模組名稱: MergeSheetsToSummary
'功能說明: 將活頁簿內多個工作表的資料合併至一個彙總工作表
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/7
'
'*************************************************************************************

' 測試用入口（先建立範例工作表再合併）
Sub TestMergeSheetsToSummary()
    Call CreateSampleSheets
    Call MergeAllSheetsToSummary
End Sub

' 建立三個季度工作表作為範例來源
Private Sub CreateSampleSheets()
    Dim sheetNames(1 To 3) As String
    Dim i As Integer
    Dim ws As Worksheet

    sheetNames(1) = "第一季"
    sheetNames(2) = "第二季"
    sheetNames(3) = "第三季"

    For i = 1 To 3
        Set ws = GetOrCreateSheet(ThisWorkbook, sheetNames(i))
        Call FillQuarterData(ws, sheetNames(i))
    Next i
End Sub

' 合併所有非彙總工作表到「合併彙總」工作表
Sub MergeAllSheetsToSummary()
    Dim wsTarget As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim colCount As Long
    Dim isFirst As Boolean

    Set wsTarget = GetOrCreateSheet(ThisWorkbook, "合併彙總")
    targetRow = 1
    isFirst = True

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "合併彙總" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

            If lastRow >= 1 Then
                If isFirst Then
                    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, colCount)).Copy _
                        Destination:=wsTarget.Cells(targetRow, 1)
                    targetRow = targetRow + lastRow
                    isFirst = False
                Else
                    If lastRow >= 2 Then
                        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, colCount)).Copy _
                            Destination:=wsTarget.Cells(targetRow, 1)
                        targetRow = targetRow + lastRow - 1
                    End If
                End If
            End If
        End If
    Next ws

    wsTarget.Columns.AutoFit
    Application.ScreenUpdating = True
    wsTarget.Activate
    MsgBox "跨工作表資料合併完成！", vbInformation, "完成"
End Sub

' 填入季度銷售資料
Private Sub FillQuarterData(ByVal ws As Worksheet, ByVal quarterName As String)
    Dim startRow As Integer
    startRow = 2

    ws.Range("A1").Value = "季度"
    ws.Range("B1").Value = "產品"
    ws.Range("C1").Value = "銷售額"

    ws.Cells(startRow, 1).Value = quarterName
    ws.Cells(startRow, 2).Value = "產品A"
    ws.Cells(startRow, 3).Value = 35000

    ws.Cells(startRow + 1, 1).Value = quarterName
    ws.Cells(startRow + 1, 2).Value = "產品B"
    ws.Cells(startRow + 1, 3).Value = 48000

    ws.Cells(startRow + 2, 1).Value = quarterName
    ws.Cells(startRow + 2, 2).Value = "產品C"
    ws.Cells(startRow + 2, 3).Value = 27000

    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
