Option Explicit
Attribute VB_Name = "SplitSheetByBlankRow"
'*************************************************************************************
'模組名稱: 依空白列分割工作表
'功能說明: 將工作表中以空白列為分隔的資料區塊，各自複製到新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub TestSplitSheetByBlankRow()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "請先選取要分割的工作表。", vbExclamation, "提示"
        Exit Sub
    End If
    Call SplitSheetByBlankRow(ws)
End Sub

Sub SplitSheetByBlankRow(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim i As Long
    Dim blockStart As Long
    Dim blockIndex As Integer
    Dim wsNew As Worksheet
    Dim blockRange As Range

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    blockStart = 1
    blockIndex = 0

    For i = 1 To lastRow + 1
        Dim isBlank As Boolean
        If i > lastRow Then
            isBlank = True
        Else
            isBlank = (Trim(CStr(ws.Cells(i, 1).Value)) = "")
        End If

        If isBlank Then
            If i > blockStart Then
                ' 找到一個非空白區塊
                blockIndex = blockIndex + 1
                Set wsNew = ws.Parent.Worksheets.Add(After:=ws.Parent.Worksheets(ws.Parent.Worksheets.Count))
                wsNew.Name = ws.Name & "_區塊" & blockIndex

                Set blockRange = ws.Range(ws.Rows(blockStart), ws.Rows(i - 1))
                blockRange.Copy Destination:=wsNew.Range("A1")
                wsNew.Columns.AutoFit
            End If
            blockStart = i + 1
        End If
    Next i

    If blockIndex = 0 Then
        MsgBox "未偵測到以空白列分隔的資料區塊。", vbInformation, "提示"
    Else
        MsgBox "已分割成 " & blockIndex & " 個工作表。", vbInformation, "完成"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "分割時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
