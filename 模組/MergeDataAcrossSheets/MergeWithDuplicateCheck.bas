Attribute VB_Name = "MergeWithDuplicateCheck"
Option Explicit
'*************************************************************************************
'模組名稱: 跨表合併並檢查重複
'功能說明: 合併活頁簿內所有工作表資料，並依關鍵欄位去除重複列
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub MergeWithDuplicateCheck()
    Dim wsDest As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim destRow As Long
    Dim i As Long
    Dim keyCol As Long
    Dim keyColStr As String
    Dim keyDict As Object
    Dim keyVal As String
    Dim isFirstSheet As Boolean
    Dim startRow As Long

    keyColStr = InputBox("請輸入關鍵欄號（用於去除重複，例如：1）：", "設定關鍵欄", "1")
    If keyColStr = "" Then Exit Sub
    If Not IsNumeric(keyColStr) Then
        MsgBox "請輸入有效的欄號。", vbExclamation, "錯誤"
        Exit Sub
    End If
    keyCol = CLng(keyColStr)

    Set keyDict = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("合併去重結果").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsDest = ThisWorkbook.Worksheets.Add
    wsDest.Name = "合併去重結果"

    destRow = 1
    isFirstSheet = True

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "合併去重結果" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

            If lastRow >= 1 And lastCol >= keyCol Then
                If isFirstSheet Then
                    ws.Rows(1).Copy Destination:=wsDest.Rows(destRow)
                    destRow = destRow + 1
                    startRow = 2
                    isFirstSheet = False
                Else
                    startRow = 2
                End If

                For i = startRow To lastRow
                    keyVal = CStr(ws.Cells(i, keyCol).Value)
                    If keyVal <> "" Then
                        If Not keyDict.Exists(keyVal) Then
                            ws.Rows(i).Copy Destination:=wsDest.Rows(destRow)
                            keyDict.Add keyVal, True
                            destRow = destRow + 1
                        End If
                    End If
                Next i
            End If
        End If
    Next ws

    wsDest.Columns.AutoFit
    MsgBox "合併去重完成！共 " & (destRow - 2) & " 列不重複資料。", vbInformation, "完成"
End Sub
