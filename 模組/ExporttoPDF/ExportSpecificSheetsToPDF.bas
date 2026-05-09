Attribute VB_Name = "ExportSpecificSheetsToPDF"
Option Explicit

' ============================================================
' 範例：將使用者指定的工作表名稱清單合併匯出成一份 PDF
' 功能：以逗號分隔輸入工作表名稱，合併列印為單一 PDF 檔
' ============================================================

Sub ExportSpecificSheetsToPDF()
    Dim inputNames  As String
    Dim nameArr()   As String
    Dim wsNames()   As String
    Dim pdfPath     As String
    Dim i           As Integer
    Dim validCount  As Integer
    Dim sheetName   As String

    inputNames = InputBox("請輸入要匯出的工作表名稱（以逗號分隔）：", "指定工作表匯出 PDF")
    If Trim(inputNames) = "" Then
        MsgBox "未輸入任何工作表名稱，操作取消。", vbInformation, "提示"
        Exit Sub
    End If

    nameArr = Split(inputNames, ",")
    ReDim wsNames(0 To UBound(nameArr))
    validCount = 0

    For i = 0 To UBound(nameArr)
        sheetName = Trim(nameArr(i))
        If SheetExists(sheetName) Then
            wsNames(validCount) = sheetName
            validCount = validCount + 1
        Else
            MsgBox "找不到工作表：" & sheetName & "，將略過。", vbExclamation, "警告"
        End If
    Next i

    If validCount = 0 Then
        MsgBox "沒有有效的工作表可匯出。", vbExclamation, "警告"
        Exit Sub
    End If

    ReDim Preserve wsNames(0 To validCount - 1)

    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "請選擇 PDF 儲存位置"
        .InitialFileName = "SpecificSheets.pdf"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        pdfPath = .SelectedItems(1)
    End With

    If LCase(Right(pdfPath, 4)) <> ".pdf" Then pdfPath = pdfPath & ".pdf"

    On Error GoTo ErrHandler
    ThisWorkbook.Sheets(wsNames).Select
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ThisWorkbook.Sheets(1).Select
    MsgBox "指定工作表已匯出為 PDF：" & vbCrLf & pdfPath, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "匯出失敗：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = (Not ws Is Nothing)
    On Error GoTo 0
End Function
