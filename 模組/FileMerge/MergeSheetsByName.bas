Attribute VB_Name = "MergeSheetsByName"
Option Explicit

' ============================================================
' 範例：從多個活頁簿中，找出相同名稱的工作表並合併其資料
' 功能：使用者指定工作表名稱，程式自動搜尋資料夾內所有
'       .xlsx 活頁簿，找到同名工作表後逐一貼入目標工作表
' ============================================================

Sub MergeSheetsBySheetName()
    On Error GoTo ErrHandler

    Dim strFolder       As String
    Dim strSheetName    As String
    Dim strFile         As String
    Dim wbSrc           As Workbook
    Dim wsSrc           As Worksheet
    Dim wsDest          As Worksheet
    Dim lngLastRow      As Long
    Dim lngDestRow      As Long
    Dim blnFirst        As Boolean
    Dim lngFilesFound   As Long

    ' --- 取得工作表名稱 ---
    strSheetName = InputBox("請輸入要合併的工作表名稱：", "指定工作表名稱", "Sheet1")
    If Len(Trim(strSheetName)) = 0 Then
        MsgBox "未輸入工作表名稱，已取消。", vbInformation, "提示"
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
    wsDest.Name = "Merged_" & Left(strSheetName, 20) & "_" & Format(Now(), "mmddHHmm")
    lngDestRow = 1
    blnFirst = True
    lngFilesFound = 0

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
                Set wsSrc = Nothing
                On Error Resume Next
                Set wsSrc = wbSrc.Worksheets(strSheetName)
                On Error GoTo ErrHandler

                If Not wsSrc Is Nothing Then
                    lngLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
                    If lngLastRow >= 1 Then
                        Dim lngStartRow As Long
                        If blnFirst Then
                            lngStartRow = 1
                            blnFirst = False
                        Else
                            lngStartRow = 2  ' 跳過標題
                        End If
                        If lngLastRow >= lngStartRow Then
                            wsSrc.Rows(lngStartRow & ":" & lngLastRow).Copy _
                                Destination:=wsDest.Cells(lngDestRow, 1)
                            lngDestRow = lngDestRow + (lngLastRow - lngStartRow + 1)
                        End If
                        lngFilesFound = lngFilesFound + 1
                    End If
                End If
                wbSrc.Close SaveChanges:=False
            End If
        End If
        strFile = Dir()
    Loop

    wsDest.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    If lngFilesFound = 0 Then
        MsgBox "找不到任何包含工作表「" & strSheetName & "」的活頁簿。", vbExclamation, "警告"
    Else
        MsgBox "依工作表名稱合併完成！" & vbCrLf & _
               "共找到 " & lngFilesFound & " 個檔案，" & _
               "合計 " & lngDestRow - 1 & " 列資料。", vbInformation, "完成"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併過程發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub