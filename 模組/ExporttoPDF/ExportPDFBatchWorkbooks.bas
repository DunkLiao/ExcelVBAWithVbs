Attribute VB_Name = "ExportPDFBatchWorkbooks"
Option Explicit

' ============================================================
' 範例：批次開啟資料夾中的 Excel 檔案，逐一將第一個工作表匯出為 PDF
' 功能：自動處理來源資料夾中所有 .xlsx / .xls 檔案
' ============================================================

Sub ExportPDFBatchWorkbooks()
    Dim srcFolder   As String
    Dim outFolder   As String
    Dim fileName    As String
    Dim wbPath      As String
    Dim pdfPath     As String
    Dim wb          As Workbook
    Dim totalCount  As Integer
    Dim e           As Integer
    Dim extensions(1) As String

    ' 選擇來源資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 Excel 檔案的來源資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        srcFolder = .SelectedItems(1)
    End With

    ' 選擇輸出資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        outFolder = .SelectedItems(1)
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False
    totalCount = 0

    extensions(0) = "*.xlsx"
    extensions(1) = "*.xls"

    For e = 0 To 1
        fileName = Dir(srcFolder & "\" & extensions(e))
        Do While fileName <> ""
            ' 略過 Excel 暫存檔（~ 開頭）
            If Left(fileName, 1) <> "~" Then
                wbPath  = srcFolder & "\" & fileName
                pdfPath = outFolder & "\" & Left(fileName, InStrRev(fileName, ".") - 1) & ".pdf"
                On Error Resume Next
                Set wb = Workbooks.Open(Filename:=wbPath, ReadOnly:=True)
                If Err.Number = 0 Then
                    wb.Sheets(1).ExportAsFixedFormat _
                        Type:=xlTypePDF, _
                        Filename:=pdfPath, _
                        Quality:=xlQualityStandard, _
                        IncludeDocProperties:=True, _
                        IgnorePrintAreas:=False, _
                        OpenAfterPublish:=False
                    wb.Close SaveChanges:=False
                    totalCount = totalCount + 1
                End If
                Err.Clear
                On Error GoTo 0
            End If
            fileName = Dir
        Loop
    Next e

    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True

    MsgBox "批次匯出完成！共匯出 " & totalCount & " 個 PDF 檔案。", vbInformation, "完成"
End Sub
