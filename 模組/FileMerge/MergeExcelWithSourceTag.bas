Attribute VB_Name = "MergeExcelWithSourceTag"
Option Explicit

' ============================================================
' 範例：合併多個 Excel 檔案，並在每列前加入來源檔案名稱標記
' 功能：搜尋資料夾內所有 .xlsx，合併第一個工作表資料，
'       同時在最左側自動插入「來源檔案」欄，記錄資料出處
' ============================================================

Sub MergeExcelWithSourceTag()
    On Error GoTo ErrHandler

    Dim strFolder       As String
    Dim strFile         As String
    Dim wbSrc           As Workbook
    Dim wsSrc           As Worksheet
    Dim wsDest          As Worksheet
    Dim lngLastRow      As Long
    Dim lngLastCol      As Long
    Dim lngDestRow      As Long
    Dim blnFirst        As Boolean
    Dim lngFilesCount   As Long
    Dim lngRow          As Long

    ' --- 選擇資料夾 ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 Excel 檔案的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    ' --- 建立目標工作表 ---
    Set wsDest = ThisWorkbook.Worksheets.Add
    wsDest.Name = "TaggedMerge_" & Format(Now(), "mmddHHmm")
    lngDestRow = 1
    blnFirst = True
    lngFilesCount = 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    strFile = Dir(strFolder & "*.xlsx")
    If strFile = "" Then
        MsgBox "找不到任何 .xlsx 檔案。", vbExclamation, "警告"
        GoTo CleanUp
    End If

    Do While strFile <> ""
        If StrComp(strFile, ThisWorkbook.Name, vbTextCompare) <> 0 Then
            Set wbSrc = Nothing
            On Error Resume Next
            Set wbSrc = Workbooks.Open(Filename:=strFolder & strFile, ReadOnly:=True)
            On Error GoTo ErrHandler

            If Not wbSrc Is Nothing Then
                Set wsSrc = wbSrc.Sheets(1)
                lngLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
                lngLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

                If lngLastRow >= 1 Then
                    If blnFirst Then
                        ' 第一個檔案：先寫入標題列，B 欄起為原始標題
                        wsDest.Cells(lngDestRow, 1).Value = "來源檔案"
                        wsSrc.Cells(1, 1).Resize(1, lngLastCol).Copy _
                            Destination:=wsDest.Cells(lngDestRow, 2)
                        lngDestRow = lngDestRow + 1
                        blnFirst = False
                    End If

                    ' 從第 2 列（跳過標題）開始複製資料
                    Dim lngSrcRow As Long
                    For lngSrcRow = 2 To lngLastRow
                        wsDest.Cells(lngDestRow, 1).Value = strFile  ' 來源標記
                        wsSrc.Cells(lngSrcRow, 1).Resize(1, lngLastCol).Copy _
                            Destination:=wsDest.Cells(lngDestRow, 2)
                        lngDestRow = lngDestRow + 1
                    Next lngSrcRow
                    lngFilesCount = lngFilesCount + 1
                End If

                wbSrc.Close SaveChanges:=False
            End If
        End If
        strFile = Dir()
    Loop

    wsDest.Columns.AutoFit

CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    If lngFilesCount > 0 Then
        MsgBox "含來源標記合併完成！" & vbCrLf & _
               "共合併 " & lngFilesCount & " 個檔案，" & _
               "合計 " & lngDestRow - 2 & " 列資料。", vbInformation, "完成"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併過程發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub