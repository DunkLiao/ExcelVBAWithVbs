Attribute VB_Name = "ExportPDFWithBatesNumbering"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithBatesNumbering
'功能說明: 將工作表匯出為 PDF 並在每頁頁尾加上 Bates 流水編號
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestExportPDFWithBatesNumbering()
    Call ExportPDFWithBates
End Sub

Sub ExportPDFWithBates()
    Dim ws As Worksheet
    Dim pdfPath As String
    Dim batesPrefix As String
    Dim startNumber As Long
    Dim totalPages As Long
    Dim i As Long
    Dim defaultName As String
    
    On Error GoTo ErrHandler
    
    Set ws = ActiveSheet
    
    ' 取得 Bates 編號設定
    batesPrefix = InputBox("請輸入 Bates 編號前置碼（例如：DEF-）", "Bates 編號設定", "DOC-")
    If batesPrefix = "" Then
        MsgBox "已取消操作", vbInformation, "取消"
        Exit Sub
    End If
    
    Dim startNumStr As String
    startNumStr = InputBox("請輸入起始流水號（例如：1）", "Bates 編號設定", "1")
    If startNumStr = "" Then
        MsgBox "已取消操作", vbInformation, "取消"
        Exit Sub
    End If
    startNumber = CLng(startNumStr)
    
    ' 選擇 PDF 存檔位置
    defaultName = ws.Name & "_Bates.pdf"
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "請選擇 PDF 存檔位置"
        .InitialFileName = defaultName
        .FilterIndex = 1
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        pdfPath = .SelectedItems(1)
    End With
    
    If LCase(Right(pdfPath, 4)) <> ".pdf" Then
        pdfPath = pdfPath & ".pdf"
    End If
    
    ' 估算總頁數
    Application.ScreenUpdating = False
    
    totalPages = CInt(Application.ExecuteExcel4Macro("GET.DOCUMENT(50)"))
    If totalPages = 0 Then totalPages = 1
    
    ' 設定頁尾 Bates 編號格式
    Dim footerText As String
    footerText = "&8&F  " & batesPrefix
    
    ' 設定頁尾為自訂文字
    With ws.PageSetup
        .CenterFooter = footerText
        .RightFooter = "第 &P 頁，共 &N 頁"
        .LeftFooter = "機密文件 - 僅供內部使用"
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    
    ' 匯出為 PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    ' 還原頁尾
    ws.PageSetup.CenterFooter = ""
    ws.PageSetup.RightFooter = ""
    ws.PageSetup.LeftFooter = ""
    
    Application.ScreenUpdating = True
    
    MsgBox "PDF 匯出完成！" & vbCrLf & _
           "Bates 前置碼：" & batesPrefix & vbCrLf & _
           "起始編號：" & Format(startNumber, "00000") & vbCrLf & _
           "存檔路徑：" & pdfPath, vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    ws.PageSetup.CenterFooter = ""
    ws.PageSetup.RightFooter = ""
    ws.PageSetup.LeftFooter = ""
    MsgBox "匯出 PDF 時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
