Attribute VB_Name = "FilterByDuplicates"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByDuplicates
'功能說明: 篩選指定欄位中的重複值或唯一值，並將結果輸出至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

' 篩選重複值（出現超過一次的列）
Sub FilterDuplicateRows()
    Call FilterByDuplicateMode(True)
End Sub

' 篩選唯一值（只出現一次的列）
Sub FilterUniqueOnlyRows()
    Call FilterByDuplicateMode(False)
End Sub

' 依重複/唯一模式篩選並輸出結果
' keepDuplicates: True=保留重複值, False=保留唯一值
Private Sub FilterByDuplicateMode(ByVal keepDuplicates As Boolean)
    Dim srcWs       As Worksheet
    Dim dstWs       As Worksheet
    Dim lastRow     As Long
    Dim lastCol     As Long
    Dim colIdx      As Long
    Dim i           As Long
    Dim destRow     As Long
    Dim dict        As Object
    Dim key         As String
    Dim modeLabel   As String

    Set srcWs = ActiveSheet
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "資料不足，請確認工作表有標題列與資料列。", vbExclamation, "提示"
        Exit Sub
    End If

    lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

    Dim colInput As String
    colInput = InputBox("請輸入要判斷重複的欄號（例如：1 代表 A 欄）：", "設定判斷欄位", "1")
    If colInput = "" Then Exit Sub
    colIdx = CLng(colInput)

    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        key = CStr(srcWs.Cells(i, colIdx).Value)
        If dict.exists(key) Then
            dict(key) = dict(key) + 1
        Else
            dict.Add key, 1
        End If
    Next i

    If keepDuplicates Then
        modeLabel = "重複值"
    Else
        modeLabel = "唯一值"
    End If

    Dim dstName As String
    dstName = Left(srcWs.Name & "_" & modeLabel, 31)

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(dstName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set dstWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    dstWs.Name = dstName

    srcWs.Rows(1).Copy Destination:=dstWs.Rows(1)
    destRow = 2

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        key = CStr(srcWs.Cells(i, colIdx).Value)
        Dim cnt As Long
        cnt = CLng(dict(key))

        If keepDuplicates And cnt > 1 Then
            srcWs.Rows(i).Copy Destination:=dstWs.Rows(destRow)
            destRow = destRow + 1
        ElseIf Not keepDuplicates And cnt = 1 Then
            srcWs.Rows(i).Copy Destination:=dstWs.Rows(destRow)
            destRow = destRow + 1
        End If
    Next i

    dstWs.Columns.AutoFit
    Application.ScreenUpdating = True

    MsgBox "篩選完成！已將「" & modeLabel & "」輸出至工作表「" & dstName & "」" & Chr(10) & _
        "共篩選出 " & (destRow - 2) & " 列資料。", vbInformation, "完成"
End Sub
