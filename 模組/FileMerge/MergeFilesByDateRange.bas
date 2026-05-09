Attribute VB_Name = "MergeFilesByDateRange"
Option Explicit

' ============================================================
' 範例：依檔案修改日期範圍篩選，合併符合條件的 Excel 檔案
' 功能：使用者輸入起始與結束日期，程式僅合併在此日期範圍
'       內被修改的 .xlsx 檔案之第一個工作表資料
' ============================================================

Sub MergeExcelFilesByDateRange()
    On Error GoTo ErrHandler

    Dim strFolder       As String
    Dim strFile         As String
    Dim strStartDate    As String
    Dim strEndDate      As String
    Dim dtStart         As Date
    Dim dtEnd           As Date
    Dim dtFileDate      As Date
    Dim wbSrc           As Workbook
    Dim wsSrc           As Worksheet
    Dim wsDest          As Worksheet
    Dim lngLastRow      As Long
    Dim lngLastCol      As Long
    Dim lngDestRow      As Long
    Dim blnFirst        As Boolean
    Dim lngFilesCount   As Long

    ' --- 取得日期範圍 ---
    strStartDate = InputBox("請輸入起始日期（格式：YYYY/MM/DD）：", "起始日期", Format(Date - 30, "YYYY/MM/DD"))
    If Not IsDate(strStartDate) Then
        MsgBox "起始日期格式錯誤，已取消。", vbExclamation, "警告"
        Exit Sub
    End If
    dtStart = CDate(strStartDate)

    strEndDate = InputBox("請輸入結束日期（格式：YYYY/MM/DD）：", "結束日期", Format(Date, "YYYY/MM/DD"))
    If Not IsDate(strEndDate) Then
        MsgBox "結束日期格式錯誤，已取消。", vbExclamation, "警告"
        Exit Sub
    End If
    dtEnd = CDate(strEndDate)

    If dtStart > dtEnd Then
        MsgBox "起始日期不可晚於結束日期，已取消。", vbExclamation, "警告"
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
    wsDest.Name = "DateFiltered_" & Format(Now(), "mmddHHmm")
    lngDestRow = 1
    blnFirst = True
    lngFilesCount = 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    strFile = Dir(strFolder & "*.xlsx")
    Do While strFile <> ""
        If StrComp(strFile, ThisWorkbook.Name, vbTextCompare) <> 0 Then
            ' 取得檔案修改日期
            On Error Resume Next
            dtFileDate = CDate(FileDateTime(strFolder & strFile))
            On Error GoTo ErrHandler

            ' 判斷是否在日期範圍內
            If dtFileDate >= dtStart And dtFileDate <= dtEnd + 1 Then
                Set wbSrc = Nothing
                On Error Resume Next
                Set wbSrc = Workbooks.Open(Filename:=strFolder & strFile, ReadOnly:=True)
                On Error GoTo ErrHandler

                If Not wbSrc Is Nothing Then
                    Set wsSrc = wbSrc.Sheets(1)
                    lngLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
                    lngLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

                    If lngLastRow >= 1 Then
                        Dim lngStartRow As Long
                        If blnFirst Then
                            lngStartRow = 1
                            blnFirst = False
                        Else
                            lngStartRow = 2
                        End If
                        If lngLastRow >= lngStartRow Then
                            wsSrc.Cells(lngStartRow, 1).Resize(lngLastRow - lngStartRow + 1, lngLastCol).Copy _
                                Destination:=wsDest.Cells(lngDestRow, 1)
                            lngDestRow = lngDestRow + (lngLastRow - lngStartRow + 1)
                        End If
                        lngFilesCount = lngFilesCount + 1
                    End If
                    wbSrc.Close SaveChanges:=False
                End If
            End If
        End If
        strFile = Dir()
    Loop

    wsDest.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "依日期範圍合併完成！" & vbCrLf & _
           "日期範圍：" & Format(dtStart, "YYYY/MM/DD") & " ~ " & Format(dtEnd, "YYYY/MM/DD") & vbCrLf & _
           "共合併 " & lngFilesCount & " 個檔案，" & _
           "合計 " & lngDestRow - 1 & " 列資料。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併過程發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub