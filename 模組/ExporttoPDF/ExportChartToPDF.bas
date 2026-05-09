Attribute VB_Name = "ExportChartToPDF"
Option Explicit

' ============================================================
' 範例：將工作表中的圖表匯出為 PDF
' 功能：支援匯出目前選取圖表、第一個圖表，或逐一匯出全部圖表
' ============================================================

' 匯出目前已選取的圖表為 PDF
Sub ExportSelectedChartToPDF()
    Dim pdfPath As String
    Dim cht     As Chart

    If TypeName(ActiveChart) = "Chart" Then
        Set cht = ActiveChart
    ElseIf TypeName(Selection) = "ChartObject" Then
        Set cht = Selection.Chart
    Else
        MsgBox "請先點選要匯出的圖表。", vbExclamation, "提示"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "請選擇圖表 PDF 儲存位置"
        .InitialFileName = "Chart.pdf"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        pdfPath = .SelectedItems(1)
    End With

    If LCase(Right(pdfPath, 4)) <> ".pdf" Then pdfPath = pdfPath & ".pdf"

    On Error GoTo ErrHandler
    cht.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "圖表已成功匯出為 PDF：" & vbCrLf & pdfPath, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "圖表匯出失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

' 匯出工作表中第一個圖表為 PDF（不需手動選取）
Sub ExportFirstChartToPDF()
    Dim ws          As Worksheet
    Dim folderPath  As String
    Dim pdfPath     As String

    Set ws = ActiveSheet

    If ws.ChartObjects.Count = 0 Then
        MsgBox "目前工作表沒有任何圖表。", vbExclamation, "提示"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    pdfPath = folderPath & "\" & ws.ChartObjects(1).Name & ".pdf"

    On Error GoTo ErrHandler
    ws.ChartObjects(1).Chart.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "圖表已匯出為 PDF：" & vbCrLf & pdfPath, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "匯出失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

' 逐一匯出工作表中所有圖表為個別 PDF
Sub ExportAllChartsToPDF()
    Dim ws          As Worksheet
    Dim folderPath  As String
    Dim pdfPath     As String
    Dim i           As Integer
    Dim totalCount  As Integer

    Set ws = ActiveSheet

    If ws.ChartObjects.Count = 0 Then
        MsgBox "目前工作表沒有任何圖表。", vbExclamation, "提示"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇 PDF 輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    totalCount = 0
    For i = 1 To ws.ChartObjects.Count
        pdfPath = folderPath & "\" & ws.ChartObjects(i).Name & ".pdf"
        On Error Resume Next
        ws.ChartObjects(i).Chart.ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=pdfPath, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=False, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
        If Err.Number = 0 Then totalCount = totalCount + 1
        Err.Clear
        On Error GoTo 0
    Next i

    MsgBox "已匯出 " & totalCount & " 個圖表為 PDF。", vbInformation, "完成"
End Sub
