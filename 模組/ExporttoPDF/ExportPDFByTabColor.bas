Attribute VB_Name = "ExportPDFByTabColor"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFByTabColor
'功能說明: 依工作表標籤顏色，將特定顏色的工作表匯出為獨立 PDF 檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub ExportPDFByTabColor()
    Dim wb          As Workbook
    Dim ws          As Worksheet
    Dim savePath    As String
    Dim pdfPath     As String
    Dim targetColor As Long
    Dim exportCount As Integer

    Set wb = ThisWorkbook

    ' 目標顏色：紅色標籤（RGB 255,0,0）
    ' 可依需求修改此顏色值
    targetColor = RGB(255, 0, 0)

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=wb.Path & "\", _
        FileFilter:="PDF 檔案 (*.pdf),*.pdf", _
        Title:="選擇 PDF 儲存資料夾（選任一檔案以確認資料夾）")

    If savePath = "False" Then
        MsgBox "已取消操作。", vbExclamation, "取消"
        Exit Sub
    End If

    savePath = Left(savePath, InStrRev(savePath, "\"))

    exportCount = 0
    Application.ScreenUpdating = False

    For Each ws In wb.Worksheets
        If ws.Tab.Color = targetColor Then
            pdfPath = savePath & ws.Name & ".pdf"
            ws.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=pdfPath, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=False
            exportCount = exportCount + 1
        End If
    Next ws

    Application.ScreenUpdating = True

    If exportCount = 0 Then
        MsgBox "找不到標籤顏色為紅色的工作表，請先設定工作表標籤顏色後再執行。", _
            vbExclamation, "未找到符合條件的工作表"
    Else
        MsgBox "已成功匯出 " & exportCount & " 個工作表為 PDF，" & _
            "儲存位置：" & savePath, vbInformation, "完成"
    End If
End Sub
