Attribute VB_Name = "MergeWithProgressIndicator"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithProgressIndicator
'功能說明: 合併所有工作表資料，並以狀態列即時顯示合併進度指示
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestMergeWithProgressIndicator()
    Call MergeAllSheetsWithProgress
End Sub

' 合併所有工作表並顯示進度指示
Sub MergeAllSheetsWithProgress()
    Dim wb As Workbook
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim lngDestRow As Long
    Dim lngLastRow As Long
    Dim lngStart As Long
    Dim blnFirst As Boolean
    Dim intTotal As Integer
    Dim intCurrent As Integer
    Dim sDestName As String

    On Error GoTo ErrHandler
    Set wb = ThisWorkbook
    sDestName = "合併結果"

    intTotal = 0
    For Each wsSrc In wb.Worksheets
        If wsSrc.Name <> sDestName Then
            intTotal = intTotal + 1
        End If
    Next wsSrc

    If intTotal = 0 Then
        MsgBox "找不到可合併的工作表。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    On Error Resume Next
    Set wsDest = wb.Worksheets(sDestName)
    On Error GoTo ErrHandler

    If wsDest Is Nothing Then
        Set wsDest = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsDest.Name = sDestName
    Else
        wsDest.Cells.Clear
    End If

    lngDestRow = 1
    blnFirst = True
    intCurrent = 0

    For Each wsSrc In wb.Worksheets
        If wsSrc.Name <> sDestName Then
            intCurrent = intCurrent + 1
            Application.StatusBar = "正在合併工作表 [" & intCurrent & "/" & intTotal & "]：" & wsSrc.Name

            lngLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

            If blnFirst Then
                lngStart = 1
                blnFirst = False
            Else
                lngStart = 2
            End If

            If lngLastRow >= lngStart Then
                wsSrc.Rows(lngStart & ":" & lngLastRow).Copy _
                    Destination:=wsDest.Cells(lngDestRow, 1)
                lngDestRow = lngDestRow + (lngLastRow - lngStart + 1)
            End If
        End If
    Next wsSrc

    wsDest.Columns.AutoFit
    wsDest.Activate
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "合併完成！共合併 " & intTotal & " 個工作表，資料已輸出至「" & sDestName & "」", _
        vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub
