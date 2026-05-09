Attribute VB_Name = "MergeFilteredRows"
Option Explicit

' ============================================================
' 範例：從多個 Excel 檔案中，僅合併符合特定條件的資料列
' 功能：使用者指定要篩選的欄號與關鍵字，程式自動找出
'       資料夾內所有 .xlsx 活頁簿符合條件的列並合併
' ============================================================

Sub MergeFilteredRowsFromFolder()
    On Error GoTo ErrHandler

    Dim strFolder       As String
    Dim strFile         As String
    Dim strKeyword      As String
    Dim strInput        As String
    Dim lngFilterCol    As Long
    Dim wbSrc           As Workbook
    Dim wsSrc           As Worksheet
    Dim wsDest          As Worksheet
    Dim lngSrcLastRow   As Long
    Dim lngSrcLastCol   As Long
    Dim lngDestRow      As Long
    Dim lngRow          As Long
    Dim blnFirst        As Boolean
    Dim lngMatchCount   As Long

    ' --- 取得篩選欄號 ---
    strInput = InputBox("請輸入要篩選的欄號（如：2 代表 B 欄）：", "篩選欄號", "2")
    If Not IsNumeric(strInput) Or CLng(strInput) < 1 Then
        MsgBox "欄號輸入無效，已取消。", vbExclamation, "警告"
        Exit Sub
    End If
    lngFilterCol = CLng(strInput)

    ' --- 取得篩選關鍵字 ---
    strKeyword = InputBox("請輸入篩選關鍵字（完全符合）：", "篩選條件", "")
    If Len(Trim(strKeyword)) = 0 Then
        MsgBox "未輸入關鍵字，已取消。", vbInformation, "提示"
        Exit Sub
    End If

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
    wsDest.Name = "Filtered_" & Format(Now(), "mmddHHmm")
    lngDestRow = 1
    blnFirst = True
    lngMatchCount = 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    strFile = Dir(strFolder & "*.xlsx")
    Do While strFile <> ""
        If StrComp(strFile, ThisWorkbook.Name, vbTextCompare) <> 0 Then
            Set wbSrc = Nothing
            On Error Resume Next
            Set wbSrc = Workbooks.Open(Filename:=strFolder & strFile, ReadOnly:=True)
            On Error GoTo ErrHandler

            If Not wbSrc Is Nothing Then
                Set wsSrc = wbSrc.Sheets(1)
                lngSrcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
                lngSrcLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

                ' 若第一個檔案，先複製標題列
                If blnFirst And lngSrcLastRow >= 1 Then
                    wsSrc.Rows(1).Copy Destination:=wsDest.Cells(1, 1)
                    lngDestRow = 2
                    blnFirst = False
                End If

                ' 逐列檢查篩選條件（從第 2 列開始，跳過標題）
                For lngRow = 2 To lngSrcLastRow
                    If CStr(wsSrc.Cells(lngRow, lngFilterCol).Value) = strKeyword Then
                        wsSrc.Rows(lngRow).Copy Destination:=wsDest.Cells(lngDestRow, 1)
                        lngDestRow = lngDestRow + 1
                        lngMatchCount = lngMatchCount + 1
                    End If
                Next lngRow

                wbSrc.Close SaveChanges:=False
            End If
        End If
        strFile = Dir()
    Loop

    wsDest.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "條件篩選合併完成！" & vbCrLf & _
           "關鍵字：" & strKeyword & vbCrLf & _
           "共篩選出 " & lngMatchCount & " 列資料。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併過程發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub