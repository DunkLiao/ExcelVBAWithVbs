Attribute VB_Name = "MergeSheetsWithSubtotal"
Option Explicit
'*************************************************************************************
'模組名稱: MergeSheetsWithSubtotal
'功能說明: 合併多個工作表的數值資料，並在每個工作表資料末尾加上小計列
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub MergeAndAddSubtotal()
    On Error GoTo ErrHandler
    Dim wbSrc      As Workbook
    Dim wsDest     As Worksheet
    Dim ws2        As Worksheet
    Dim destRow    As Long
    Dim lastRow    As Long
    Dim lastCol    As Long
    Dim isFirst    As Boolean
    Dim sheetCount As Long
    Dim dataStart  As Long
    Dim c          As Long

    Set wbSrc = ThisWorkbook
    Set wsDest = GetOrCreateSheetSub(wbSrc, "合併含小計")
    destRow    = 1
    isFirst    = True
    sheetCount = 0

    For Each ws2 In wbSrc.Worksheets
        If ws2.Visible = xlSheetVisible And ws2.Name <> wsDest.Name Then
            lastRow = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
            lastCol = ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column
            If lastRow >= 2 And lastCol >= 1 Then
                If isFirst Then
                    ws2.Range(ws2.Cells(1, 1), ws2.Cells(1, lastCol)).Copy _
                        Destination:=wsDest.Cells(destRow, 1)
                    destRow = destRow + 1
                    isFirst = False
                End If
                dataStart = destRow
                ws2.Range(ws2.Cells(2, 1), ws2.Cells(lastRow, lastCol)).Copy _
                    Destination:=wsDest.Cells(destRow, 1)
                destRow = destRow + (lastRow - 1)
                wsDest.Cells(destRow, 1).Value = "[" & ws2.Name & "] 小計"
                For c = 2 To lastCol
                    If IsNumeric(wsDest.Cells(dataStart, c).Value) Then
                        wsDest.Cells(destRow, c).Formula = "=SUM(" & _
                            wsDest.Cells(dataStart, c).Address & ":" & _
                            wsDest.Cells(destRow - 1, c).Address & ")"
                    End If
                Next c
                With wsDest.Range(wsDest.Cells(destRow, 1), wsDest.Cells(destRow, lastCol))
                    .Font.Bold = True
                    .Interior.Color = RGB(255, 235, 156)
                End With
                destRow    = destRow + 1
                sheetCount = sheetCount + 1
            End If
        End If
    Next ws2
    wsDest.Columns.AutoFit
    wsDest.Activate
    MsgBox "已合併 " & sheetCount & " 張工作表並加上各表小計。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetSub(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetSub = ws
End Function

