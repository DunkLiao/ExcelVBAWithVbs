Attribute VB_Name = "SplitSheetToHTML"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetToHTML
'功能說明: 將 Excel 各工作表資料分別匯出為 HTML 靜態表格檔案的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點：將所有工作表匯出為 HTML
Sub TestSplitSheetToHTML()
    Dim outputDir As String
    outputDir = Environ("USERPROFILE") & "\Desktop\ExcelHTMLExport"
    Call SplitAllSheetsToHTML(ThisWorkbook, outputDir)
End Sub

' 將活頁簿每個工作表匯出為獨立 HTML 檔案
' wb: 來源活頁簿
' outputDir: 輸出資料夾路徑
Sub SplitAllSheetsToHTML(ByVal wb As Workbook, ByVal outputDir As String)
    On Error GoTo ErrorHandler

    Dim fso     As Object
    Dim ws      As Worksheet
    Dim outPath As String
    Dim count   As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outputDir) Then
        fso.CreateFolder outputDir
    End If

    count = 0
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            outPath = outputDir & "\" & CleanHTMLFileName(ws.Name) & ".html"
            Call ExportSheetAsHTML(ws, outPath, fso)
            count = count + 1
        End If
    Next ws

    MsgBox "已匯出 " & count & " 個工作表為 HTML 檔案。" & Chr(10) & _
           "輸出路徑：" & outputDir, vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "匯出 HTML 時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 將單一工作表匯出為 HTML 表格
Private Sub ExportSheetAsHTML(ByVal ws As Worksheet, ByVal filePath As String, _
                               ByVal fso As Object)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim r       As Long
    Dim c       As Long
    Dim cellVal As String
    Dim html    As String
    Dim ts      As Object

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 1 Or lastCol < 1 Then Exit Sub

    html = "<!DOCTYPE html>" & Chr(10)
    html = html & "<html><head><meta charset=""UTF-8""><title>" & ws.Name & "</title>" & Chr(10)
    html = html & "<style>table{border-collapse:collapse;font-family:Arial,sans-serif;}" & Chr(10)
    html = html & "th{background:#4472C4;color:#fff;padding:6px 10px;border:1px solid #999;}" & Chr(10)
    html = html & "td{padding:5px 10px;border:1px solid #ccc;}" & Chr(10)
    html = html & "tr:nth-child(even){background:#f2f2f2;}</style></head><body>" & Chr(10)
    html = html & "<h2>" & ws.Name & "</h2>" & Chr(10)
    html = html & "<table>" & Chr(10)

    html = html & "<tr>"
    For c = 1 To lastCol
        cellVal = HTMLEscape(CStr(ws.Cells(1, c).Value))
        html = html & "<th>" & cellVal & "</th>"
    Next c
    html = html & "</tr>" & Chr(10)

    For r = 2 To lastRow
        html = html & "<tr>"
        For c = 1 To lastCol
            cellVal = HTMLEscape(CStr(ws.Cells(r, c).Value))
            html = html & "<td>" & cellVal & "</td>"
        Next c
        html = html & "</tr>" & Chr(10)
    Next r

    html = html & "</table></body></html>"

    Set ts = fso.CreateTextFile(filePath, True, True)
    ts.Write html
    ts.Close
End Sub

' 轉義 HTML 特殊字元
Private Function HTMLEscape(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, Chr(34), "&quot;")
    HTMLEscape = s
End Function

' 清除不合法的檔名字元
Private Function CleanHTMLFileName(ByVal name As String) As String
    Dim illegalChars As String
    Dim i            As Integer
    illegalChars = "\/:*?""<>|"
    For i = 1 To Len(illegalChars)
        name = Replace(name, Mid(illegalChars, i, 1), "_")
    Next i
    CleanHTMLFileName = name
End Function
