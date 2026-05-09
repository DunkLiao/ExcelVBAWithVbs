Attribute VB_Name = "ExportPDFWithWatermark"
Option Explicit

' ============================================================
' 範例：加入文字浮水印後匯出 PDF，完成後自動移除浮水印
' 功能：在工作表中加入半透明文字方塊作為浮水印，匯出後刪除
' ============================================================

Private Const WATERMARK_NAME As String = "WatermarkShape_Temp"

Sub ExportPDFWithWatermark()
    Dim ws          As Worksheet
    Dim pdfPath     As String
    Dim waterText   As String
    Dim shp         As Shape
    Dim centerLeft  As Double
    Dim centerTop   As Double

    Set ws = ActiveSheet

    waterText = InputBox("請輸入浮水印文字：", "設定浮水印", "機密文件")
    If Trim(waterText) = "" Then
        MsgBox "未輸入浮水印文字，操作取消。", vbInformation, "提示"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "請選擇 PDF 儲存位置"
        .InitialFileName = ws.Name & "_Watermark.pdf"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        pdfPath = .SelectedItems(1)
    End With

    If LCase(Right(pdfPath, 4)) <> ".pdf" Then pdfPath = pdfPath & ".pdf"

    ' 以前 15 列的中間位置估算浮水印位置
    centerLeft = ws.Range("A1").Left + 150
    centerTop  = ws.Range("A15").Top

    On Error GoTo ErrHandler

    ' 新增浮水印文字方塊
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                   centerLeft, centerTop, 350, 80)
    With shp
        .Name = WATERMARK_NAME
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        .Rotation = 315
        With .TextFrame2.TextRange
            .Text = waterText
            With .Font
                .Size = 54
                .Bold = msoTrue
                .Fill.Visible = msoTrue
                .Fill.ForeColor.RGB = RGB(210, 210, 210)
                .Fill.Transparency = 0.3
            End With
        End With
    End With

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "含浮水印的 PDF 已匯出：" & vbCrLf & pdfPath, vbInformation, "完成"

ErrHandler:
    If Err.Number <> 0 Then
        MsgBox "操作失敗：" & Err.Description, vbCritical, "錯誤"
    End If
    RemoveWatermark ws
End Sub

Private Sub RemoveWatermark(ByVal ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = WATERMARK_NAME Then
            shp.Delete
            Exit For
        End If
    Next shp
End Sub
