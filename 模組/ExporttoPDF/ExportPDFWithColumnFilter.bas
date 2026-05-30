Attribute VB_Name = "ExportPDFWithColumnFilter"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFWithColumnFilter
'功能說明: 依使用者選擇的欄位，先篩選隱藏不需要的欄，再匯出為PDF檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestExportPDFWithColumnFilter()
    Call ExportPDFWithColumnFilter
End Sub

' 依欄位篩選後匯出PDF
Sub ExportPDFWithColumnFilter()
    Dim ws As Worksheet
    Dim sVisibleCols As String
    Dim sExportPath As String
    Dim arrCols() As String
    Dim i As Integer
    Dim j As Long
    Dim lngLastCol As Long
    Dim blnShow As Boolean

    On Error GoTo ErrHandler
    Set ws = ActiveSheet
    lngLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lngLastCol < 1 Then
        MsgBox "工作表無資料。", vbExclamation
        Exit Sub
    End If

    sVisibleCols = InputBox( _
        "請輸入要匯出的欄號（以逗號分隔），例如：1,3,5" & Chr(13) & _
        "目前工作表共有 " & lngLastCol & " 欄。", _
        "選擇匯出欄位", "1,2,3")

    If sVisibleCols = "" Then
        MsgBox "已取消操作。", vbInformation
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogSaveAs)
        .Title = "選擇PDF儲存位置"
        .InitialFileName = ws.Name & "_欄位篩選.pdf"
        .FilterIndex = 1
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation
            Exit Sub
        End If
        sExportPath = .SelectedItems(1)
    End With

    If Right(LCase(sExportPath), 4) <> ".pdf" Then
        sExportPath = sExportPath & ".pdf"
    End If

    arrCols = Split(sVisibleCols, ",")
    For j = 1 To lngLastCol
        blnShow = False
        For i = 0 To UBound(arrCols)
            If Trim(arrCols(i)) = CStr(j) Then
                blnShow = True
                Exit For
            End If
        Next i
        ws.Columns(j).Hidden = Not blnShow
    Next j

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=sExportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ws.Columns.Hidden = False
    MsgBox "PDF已匯出至：" & Chr(13) & sExportPath, vbInformation, "完成"
    Exit Sub

ErrHandler:
    On Error Resume Next
    ws.Columns.Hidden = False
    On Error GoTo 0
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub
