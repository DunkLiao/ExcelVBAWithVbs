Attribute VB_Name = "MergeAllSheetsToWorkbook"
Option Explicit

' ============================================================
' 範例：將多個活頁簿的所有工作表複製至單一新活頁簿
' 功能：搜尋指定資料夾的所有 .xlsx 活頁簿，
'       將每個活頁簿的所有工作表各別複製成獨立工作表，
'       工作表命名格式：來源檔名_工作表名
' ============================================================

Sub MergeAllSheetsIntoOneWorkbook()
    On Error GoTo ErrHandler

    Dim strFolder   As String
    Dim strFile     As String
    Dim wbSrc       As Workbook
    Dim wbDest      As Workbook
    Dim wsSrc       As Worksheet
    Dim strNewName  As String
    Dim lngSheetsCount As Long
    Dim lngFilesCount  As Long

    ' --- 選擇資料夾 ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 Excel 活頁簿的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    strFile = Dir(strFolder & "*.xlsx")
    If strFile = "" Then
        MsgBox "找不到任何 .xlsx 活頁簿。", vbExclamation, "警告"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' --- 建立新目標活頁簿 ---
    Set wbDest = Workbooks.Add
    ' 刪除預設工作表（保留最後一個以防出錯）
    Do While wbDest.Sheets.Count > 1
        wbDest.Sheets(1).Delete
    Loop

    lngSheetsCount = 0
    lngFilesCount = 0

    Do While strFile <> ""
        If StrComp(strFile, ThisWorkbook.Name, vbTextCompare) <> 0 Then
            Set wbSrc = Nothing
            On Error Resume Next
            Set wbSrc = Workbooks.Open(Filename:=strFolder & strFile, ReadOnly:=True)
            On Error GoTo ErrHandler

            If Not wbSrc Is Nothing Then
                Dim i As Long
                For i = 1 To wbSrc.Sheets.Count
                    Set wsSrc = wbSrc.Sheets(i)
                    ' 命名：前 15 字元的檔名 + _ + 工作表名
                    Dim strBaseName As String
                    strBaseName = Left(Replace(strFile, ".xlsx", ""), 15)
                    strNewName = strBaseName & "_" & Left(wsSrc.Name, 15)
                    ' 確保名稱不重複
                    Dim lngSuffix As Long
                    lngSuffix = 0
                    Dim strFinalName As String
                    strFinalName = strNewName
                    On Error Resume Next
                    Do While wbDest.Sheets(strFinalName) Is Nothing = False
                        lngSuffix = lngSuffix + 1
                        strFinalName = Left(strNewName, 28) & "_" & lngSuffix
                    Loop
                    On Error GoTo ErrHandler
                    wsSrc.Copy After:=wbDest.Sheets(wbDest.Sheets.Count)
                    wbDest.Sheets(wbDest.Sheets.Count).Name = strFinalName
                    lngSheetsCount = lngSheetsCount + 1
                Next i
                wbSrc.Close SaveChanges:=False
                lngFilesCount = lngFilesCount + 1
            End If
        End If
        strFile = Dir()
    Loop

    ' 刪除初始預留的空白工作表（Sheet1）
    On Error Resume Next
    wbDest.Sheets("Sheet1").Delete
    On Error GoTo ErrHandler

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "工作表合併至新活頁簿完成！" & vbCrLf & _
           "共處理 " & lngFilesCount & " 個檔案，" & _
           "複製 " & lngSheetsCount & " 個工作表。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "合併過程發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub