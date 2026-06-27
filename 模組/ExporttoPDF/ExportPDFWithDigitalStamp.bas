Attribute VB_Name = "ExportPDFWithDigitalStamp"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithDigitalStamp
'功能說明: 匯出PDF時自動加入日期戳記、頁碼及使用者資訊的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestExportPDFWithDigitalStamp()
    Call ExportPDFWithDigitalStamp("D:\Temp\PDFOutput")
End Sub

Sub ExportPDFWithDigitalStamp(ByVal outputPath As String)
    Dim ws As Worksheet
    Dim sheetName As String
    Dim pdfPath As String
    Dim stampText As String
    Dim fso As Object
    
    sheetName = "PDF日期戳記"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillStampSampleData(ws)
    
    stampText = "列印日期: " & Format(Now, "yyyy/mm/dd HH:mm:ss") & vbCrLf & _
                "列印人員: " & Environ("USERNAME") & vbCrLf & _
                "文件等級: 內部使用"
    
    ws.Range("F1").Value = stampText
    ws.Range("F1").Font.Size = 8
    ws.Range("F1").Font.Color = RGB(100, 100, 100)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(outputPath) Then
        fso.CreateFolder outputPath
    End If
    Set fso = Nothing
    
    With ws.PageSetup
        .PrintArea = "A1:D6"
        .LeftHeader = "公司機密文件"
        .CenterHeader = "月銷售報告"
        .RightHeader = "第 &P 頁 / 共 &N 頁"
        .LeftFooter = "&D &T"
        .RightFooter = "Confidential"
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    
    pdfPath = outputPath & "\報表_" & Format(Now, "yyyymmdd_HHmmss") & ".pdf"
    
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        fileName:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True
    
    MsgBox "PDF已匯出完成！" & vbCrLf & _
           "路徑: " & pdfPath, vbInformation, "完成"
End Sub

Private Sub FillStampSampleData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品編號"
    ws.Range("B1").Value = "產品名稱"
    ws.Range("C1").Value = "銷售數量"
    ws.Range("D1").Value = "銷售金額"
    
    ws.Range("A2").Value = "P001"
    ws.Range("B2").Value = "筆記型電腦"
    ws.Range("C2").Value = 25
    ws.Range("D2").Value = 750000
    
    ws.Range("A3").Value = "P002"
    ws.Range("B3").Value = "平板電腦"
    ws.Range("C3").Value = 40
    ws.Range("D3").Value = 480000
    
    ws.Range("A4").Value = "P003"
    ws.Range("B4").Value = "智慧手機"
    ws.Range("C4").Value = 60
    ws.Range("D4").Value = 900000
    
    ws.Range("A5").Value = "P004"
    ws.Range("B5").Value = "印表機"
    ws.Range("C5").Value = 15
    ws.Range("D5").Value = 105000
    
    ws.Range("A6").Value = ""
    ws.Range("B1:B5").Font.Bold = True
    
    ws.Columns("A:D").AutoFit
End Sub
