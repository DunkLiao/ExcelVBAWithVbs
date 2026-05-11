Attribute VB_Name = "SplitSheetToTXT"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetToTXT
'功能說明: 將 Excel 工作表依分組欄位切割並匯出為個別的 TXT 文字檔
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口：將第一個工作表依 A 欄分組匯出至桌面
Sub TestSplitSheetToTXT()
    Dim ws As Worksheet
    Dim outputFolder As String
    Set ws = ThisWorkbook.Worksheets(1)
    outputFolder = Environ("USERPROFILE") & "\Desktop\SplitTXT\"
    Call SplitSheetToTXT(ws, 1, outputFolder)
End Sub

' 將工作表資料依指定欄位分組，匯出為個別 TXT 檔
' ws           : 來源工作表
' groupCol     : 分組依據的欄號 (1=A)
' outputFolder : 輸出資料夾路徑
Sub SplitSheetToTXT( _
    ByVal ws As Worksheet, _
    ByVal groupCol As Integer, _
    ByVal outputFolder As String)

    Dim fso As Object
    Dim fileStream As Object
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r As Long
    Dim c As Integer
    Dim groupKey As String
    Dim prevGroupKey As String
    Dim headerLine As String
    Dim dataLine As String
    Dim fileCount As Integer
    Dim safeFileName As String

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(outputFolder) Then
        fso.CreateFolder outputFolder
    End If

    lastRow = ws.Cells(ws.Rows.Count, groupCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "工作表資料列數不足。", vbExclamation, "錯誤"
        Exit Sub
    End If

    headerLine = ""
    For c = 1 To lastCol
        If c > 1 Then headerLine = headerLine & vbTab
        headerLine = headerLine & CStr(ws.Cells(1, c).Value)
    Next c

    prevGroupKey = ""
    fileCount = 0

    For r = 2 To lastRow
        groupKey = Trim(CStr(ws.Cells(r, groupCol).Value))
        If groupKey = "" Then GoTo NextRow

        If groupKey <> prevGroupKey Then
            If Not fileStream Is Nothing Then
                fileStream.Close
                Set fileStream = Nothing
            End If

            safeFileName = Replace(groupKey, "/", "-")
            safeFileName = Replace(safeFileName, "\", "-")
            safeFileName = Replace(safeFileName, ":", "-")
            Set fileStream = fso.CreateTextFile( _
                outputFolder & safeFileName & ".txt", True, False)
            fileStream.WriteLine headerLine
            prevGroupKey = groupKey
            fileCount = fileCount + 1
        End If

        dataLine = ""
        For c = 1 To lastCol
            If c > 1 Then dataLine = dataLine & vbTab
            dataLine = dataLine & CStr(ws.Cells(r, c).Value)
        Next c
        fileStream.WriteLine dataLine
NextRow:
    Next r

    If Not fileStream Is Nothing Then
        fileStream.Close
        Set fileStream = Nothing
    End If

    MsgBox "分割完成！共產生 " & fileCount & " 個 TXT 檔案。" & vbCrLf & _
           "輸出路徑：" & outputFolder, vbInformation, "完成"
    Set fso = Nothing
End Sub