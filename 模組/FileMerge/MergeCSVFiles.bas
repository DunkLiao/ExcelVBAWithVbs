Attribute VB_Name = "MergeCSVFiles"
Option Explicit

' ============================================================
' 範例：合併同一資料夾下所有 CSV 檔案至單一工作表
' 功能：讀取目標資料夾內所有 .csv 並逐列合併，第一個檔案保留標題
' ============================================================
Sub MergeAllCSVFiles()
    Dim strFolder   As String
    Dim strFile     As String
    Dim wsDest      As Worksheet
    Dim wbSrc       As Workbook
    Dim wsSrc       As Worksheet
    Dim lngLastRow  As Long
    Dim lngDestRow  As Long
    Dim lngStartRow As Long
    Dim blnFirst    As Boolean

    ' --- 選擇資料夾 ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 CSV 檔案的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    ' --- 建立目標工作表 ---
    On Error GoTo ErrHandler
    Set wsDest = ThisWorkbook.Worksheets.Add
    wsDest.Name = "MergedCSV_" & Format(Now(), "mmddHHmm")
    lngDestRow = 1
    blnFirst = True

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    strFile = Dir(strFolder & "*.csv")
    If strFile = "" Then
        MsgBox "找不到任何 CSV 檔案。", vbExclamation
        GoTo CleanUp
    End If

    Do While strFile <> ""
        Set wbSrc = Workbooks.Open(Filename:=strFolder & strFile, ReadOnly:=True)
        Set wsSrc = wbSrc.Sheets(1)
        lngLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

        If blnFirst Then
            lngStartRow = 1  ' 第一個檔案含標題
            blnFirst = False
        Else
            lngStartRow = 2  ' 後續檔案跳過標題
        End If

        wsSrc.Rows(lngStartRow & ":" & lngLastRow).Copy _
            Destination:=wsDest.Cells(lngDestRow, 1)
        lngDestRow = lngDestRow + (lngLastRow - lngStartRow + 1)

        wbSrc.Close SaveChanges:=False
        strFile = Dir()
    Loop

    wsDest.Columns.AutoFit
    MsgBox "CSV 合併完成！資料已寫入工作表：" & wsDest.Name, vbInformation

CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
