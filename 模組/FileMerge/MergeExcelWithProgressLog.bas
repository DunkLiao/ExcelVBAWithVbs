Attribute VB_Name = "MergeExcelWithProgressLog"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelWithProgressLog
'功能說明: 合併多個Excel檔案，並同步記錄每個步驟的進度日誌到摘要工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestMergeExcelWithProgressLog()
    Call MergeFilesWithLog
End Sub

' 合併Excel檔案並記錄進度日誌
Sub MergeFilesWithLog()
    Dim strFolder As String
    Dim strFile As String
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim wsLog As Worksheet
    Dim lngDestRow As Long
    Dim lngLogRow As Long
    Dim blnFirst As Boolean
    Dim lngMerged As Long
    Dim lngFailed As Long
    Dim lngRows As Long
    Dim lngStart As Long
    Dim dtStart As Date

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含Excel檔案的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    dtStart = Now()

    Set wsDest = ThisWorkbook.Worksheets.Add
    wsDest.Name = "合併結果_" & Format(Now(), "mmddHHmm")
    lngDestRow = 1
    blnFirst = True

    Set wsLog = ThisWorkbook.Worksheets.Add
    wsLog.Name = "合併日誌_" & Format(Now(), "mmddHHmm")
    lngLogRow = 1

    wsLog.Range("A1").Value = "時間"
    wsLog.Range("B1").Value = "檔案名稱"
    wsLog.Range("C1").Value = "狀態"
    wsLog.Range("D1").Value = "資料列數"
    wsLog.Range("E1").Value = "備註"
    wsLog.Range("A1:E1").Font.Bold = True
    lngLogRow = 2

    lngMerged = 0
    lngFailed = 0

    strFile = Dir(strFolder & "*.xlsx")
    If strFile = "" Then strFile = Dir(strFolder & "*.xls")

    If strFile = "" Then
        MsgBox "找不到任何Excel檔案。", vbExclamation
        GoTo CleanUp
    End If

    Do While strFile <> ""
        lngRows = 0
        On Error Resume Next
        Set wbSrc = Workbooks.Open(Filename:=strFolder & strFile, ReadOnly:=True)
        If Err.Number <> 0 Then
            wsLog.Cells(lngLogRow, 1).Value = Format(Now(), "hh:mm:ss")
            wsLog.Cells(lngLogRow, 2).Value = strFile
            wsLog.Cells(lngLogRow, 3).Value = "失敗"
            wsLog.Cells(lngLogRow, 4).Value = 0
            wsLog.Cells(lngLogRow, 5).Value = Err.Description
            wsLog.Cells(lngLogRow, 3).Font.Color = RGB(255, 0, 0)
            lngLogRow = lngLogRow + 1
            lngFailed = lngFailed + 1
            Err.Clear
            strFile = Dir()
        Else
            On Error GoTo ErrHandler
            Set wsSrc = wbSrc.Sheets(1)
            lngRows = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
            If blnFirst Then
                lngStart = 1
                blnFirst = False
            Else
                lngStart = 2
            End If
            If lngRows >= lngStart Then
                wsSrc.Rows(lngStart & ":" & lngRows).Copy _
                    Destination:=wsDest.Cells(lngDestRow, 1)
                lngDestRow = lngDestRow + (lngRows - lngStart + 1)
            End If
            wbSrc.Close SaveChanges:=False
            wsLog.Cells(lngLogRow, 1).Value = Format(Now(), "hh:mm:ss")
            wsLog.Cells(lngLogRow, 2).Value = strFile
            wsLog.Cells(lngLogRow, 3).Value = "成功"
            wsLog.Cells(lngLogRow, 4).Value = lngRows - lngStart + 1
            wsLog.Cells(lngLogRow, 5).Value = ""
            wsLog.Cells(lngLogRow, 3).Font.Color = RGB(0, 128, 0)
            lngLogRow = lngLogRow + 1
            lngMerged = lngMerged + 1
            strFile = Dir()
        End If
    Loop

    wsLog.Cells(lngLogRow + 1, 1).Value = "摘要"
    wsLog.Cells(lngLogRow + 1, 2).Value = "成功：" & lngMerged & " 個，失敗：" & lngFailed & " 個"
    wsLog.Cells(lngLogRow + 1, 3).Value = "耗時：" & Format(Now() - dtStart, "hh:mm:ss")
    wsLog.Cells(lngLogRow + 1, 1).Font.Bold = True
    wsDest.Columns.AutoFit
    wsLog.Columns.AutoFit
    MsgBox "合併完成！成功：" & lngMerged & " 個，失敗：" & lngFailed & " 個", _
        vbInformation, "完成"

CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub
