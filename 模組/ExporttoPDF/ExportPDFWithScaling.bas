Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithScaling
'功能說明: 將作用中工作表匯出為 PDF，並套用使用者指定的縮放比例（Zoom%）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub ExportPDFWithScaling()
    Dim ws As Worksheet
    Dim savePath As String
    Dim zoomPct As Long
    Dim userInput As String

    On Error GoTo ErrHandler

    Set ws = ActiveSheet

    ' 詢問縮放比例
    userInput = InputBox( _
        "請輸入 PDF 匯出縮放比例（10 ~ 400，預設 80）：", _
        "設定縮放比例", "80")
    If userInput = "" Then Exit Sub

    If Not IsNumeric(userInput) Then
        MsgBox "請輸入有效的數字！", vbExclamation, "輸入錯誤"
        Exit Sub
    End If

    zoomPct = CLng(userInput)
    If zoomPct < 10 Or zoomPct > 400 Then
        MsgBox "縮放比例必須介於 10 到 400 之間！", vbExclamation, "輸入錯誤"
        Exit Sub
    End If

    ' 選擇儲存路徑
    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "儲存 PDF 檔案"
        .InitialFileName = ws.Name & "_縮放" & zoomPct & "pct.pdf"
        If .Show = False Then Exit Sub
        savePath = .SelectedItems(1)
    End With

    ' 確保副檔名為 .pdf
    If LCase(Right(savePath, 4)) <> ".pdf" Then
        savePath = savePath & ".pdf"
    End If

    ' 套用縮放比例設定
    With ws.PageSetup
        .Zoom = zoomPct
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
    End With

    ' 匯出 PDF
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "PDF 已匯出完成！" & vbNewLine & _
           "縮放比例：" & zoomPct & "%" & vbNewLine & _
           "儲存路徑：" & savePath, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
