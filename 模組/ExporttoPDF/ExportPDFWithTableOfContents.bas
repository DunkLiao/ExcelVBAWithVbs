Attribute VB_Name = "ExportPDFWithTableOfContents"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithTableOfContents
'功能說明: 將所有工作表匯出為 PDF 前，先在第一頁產生目錄頁，列出各工作表名稱
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口
Sub TestExportPDFWithTableOfContents()
    Dim savePath As String
    savePath = Environ("USERPROFILE") & "\Desktop\活頁簿目錄.pdf"
    Call ExportPDFWithTableOfContents(savePath)
End Sub

' 建立目錄工作表後匯出所有工作表為 PDF
' pdfPath: 輸出 PDF 的完整路徑
Sub ExportPDFWithTableOfContents(ByVal pdfPath As String)
    Dim tocWs As Worksheet
    Dim ws As Worksheet
    Dim sheetNames() As String
    Dim sheetCount As Integer
    Dim i As Integer
    Dim tocSheetName As String

    tocSheetName = "_目錄頁_"
    sheetCount = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> tocSheetName Then
            sheetCount = sheetCount + 1
        End If
    Next ws

    If sheetCount = 0 Then
        MsgBox "活頁簿中沒有可匯出的工作表。", vbExclamation, "錯誤"
        Exit Sub
    End If

    ReDim sheetNames(1 To sheetCount)
    i = 1
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> tocSheetName Then
            sheetNames(i) = ws.Name
            i = i + 1
        End If
    Next ws

    On Error Resume Next
    Set tocWs = ThisWorkbook.Worksheets(tocSheetName)
    On Error GoTo 0

    If tocWs Is Nothing Then
        Set tocWs = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        tocWs.Name = tocSheetName
    Else
        tocWs.Cells.Clear
        tocWs.Move Before:=ThisWorkbook.Worksheets(1)
    End If

    Call BuildTOCPage(tocWs, sheetNames, sheetCount)

    On Error GoTo ExportError
    ThisWorkbook.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    Application.DisplayAlerts = False
    tocWs.Delete
    Application.DisplayAlerts = True
    MsgBox "PDF 已匯出至：" & pdfPath, vbInformation, "完成"
    Exit Sub

ExportError:
    Application.DisplayAlerts = False
    tocWs.Delete
    Application.DisplayAlerts = True
    MsgBox "匯出 PDF 時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 建立目錄頁內容
' tocWs      : 目錄工作表
' sheetNames : 工作表名稱陣列
' sheetCount : 工作表數量
Private Sub BuildTOCPage( _
    ByVal tocWs As Worksheet, _
    ByVal sheetNames() As String, _
    ByVal sheetCount As Integer)

    Dim i As Integer

    With tocWs.Range("B2")
        .Value = "活頁簿目錄"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(31, 73, 125)
    End With

    tocWs.Range("B3").Value = "產生日期：" & Format(Now, "yyyy/mm/dd hh:mm")

    With tocWs.Range("B5:C5")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
    End With
    tocWs.Range("B5").Value = "頁次"
    tocWs.Range("C5").Value = "工作表名稱"

    For i = 1 To sheetCount
        tocWs.Range("B" & (5 + i)).Value = i
        tocWs.Range("C" & (5 + i)).Value = sheetNames(i)
    Next i

    With tocWs.Range("B5:C" & (5 + sheetCount)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    tocWs.Columns("B:C").AutoFit
End Sub