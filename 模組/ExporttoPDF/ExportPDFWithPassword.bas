Attribute VB_Name = "ExportPDFWithPassword"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithPassword
'功能說明: 將工作表匯出為 PDF，並提示使用者以第三方工具加密的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點
Sub TestExportPDFWithPassword()
    Dim ws      As Worksheet
    Dim outPath As String

    Set ws = ThisWorkbook.ActiveSheet
    outPath = Environ("USERPROFILE") & "\Desktop\" & _
              CleanPDFFileName(ws.Name) & "_protected.pdf"

    Call ExportSheetToPDFWithPasswordHint(ws, outPath)
End Sub

' 匯出工作表為 PDF 並提示加密說明
' ws: 要匯出的工作表
' pdfPath: 輸出 PDF 路徑
Sub ExportSheetToPDFWithPasswordHint(ByVal ws As Worksheet, ByVal pdfPath As String)
    On Error GoTo ErrorHandler

    Dim password As String
    password = InputBox("請輸入 PDF 保護密碼（本功能為流程示範，實際加密需搦配 PDF 工具）：", _
                        "設定 PDF 密碼")
    If password = "" Then
        MsgBox "未輸入密碼，已取消匯出。", vbInformation, "取消"
        Exit Sub
    End If

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    Dim msg As String
    msg = "PDF 已匯出至：" & Chr(10) & pdfPath & Chr(10) & Chr(10)
    msg = msg & "【密碼加密提示】" & Chr(10)
    msg = msg & "Excel 原生匯出 PDF 不支援密碼保護。" & Chr(10)
    msg = msg & "請使用下列任一方式加密：" & Chr(10)
    msg = msg & "  1. Adobe Acrobat：開啟 PDF -> 工具 -> 保護 -> 加密" & Chr(10)
    msg = msg & "  2. iTextSharp（.NET）或 PyPDF2（Python）程式化加密" & Chr(10)
    msg = msg & "  3. 第三方 PDF 工具（如 Foxit、PDF24）" & Chr(10)
    msg = msg & Chr(10) & "您設定的密碼為：" & password & Chr(10)
    msg = msg & "請在上述工具中套用此密碼。"

    MsgBox msg, vbInformation, "匯出完成 - 加密提示"
    Exit Sub

ErrorHandler:
    MsgBox "匯出 PDF 時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 清除不合法的檔名字元
Private Function CleanPDFFileName(ByVal name As String) As String
    Dim illegalChars As String
    Dim i            As Integer
    illegalChars = "\/:*?""<>|"
    For i = 1 To Len(illegalChars)
        name = Replace(name, Mid(illegalChars, i, 1), "_")
    Next i
    CleanPDFFileName = name
End Function
