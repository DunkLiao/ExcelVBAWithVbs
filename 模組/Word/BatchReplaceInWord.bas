Attribute VB_Name = "BatchReplaceInWord"
Option Explicit
'*************************************************************************************
'模組名稱: BatchReplaceInWord
'功能說明: 批次取代指定資料夾內所有 Word 文件中的特定文字
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 在 Excel 工作表 A 欄填入「尋找文字」，B 欄填入「取代文字」
'  2. 執行 BatchReplaceTextInWordFiles
'  3. 選擇要處理的 Word 文件資料夾
'  4. 程式將自動逐一開啟、取代並儲存所有 .docx 檔案
'
'注意事項:
'  - 工作表第一列為標題，資料從第二列開始
'  - 取代後原始檔案將被覆蓋，請先備份
'*************************************************************************************

'批次取代資料夾內 Word 文件文字
Sub BatchReplaceTextInWordFiles()
    Dim wsRule       As Worksheet
    Dim wdApp        As Object
    Dim wdDoc        As Object
    Dim strFolder    As String
    Dim strFile      As String
    Dim lngRow       As Long
    Dim lngLastRow   As Long
    Dim lngCount     As Long
    Dim lngFileCount As Long
    Dim arrFind()    As String
    Dim arrReplace() As String
    Dim i            As Long

    On Error GoTo ErrHandler

    Set wsRule = ActiveSheet

    '讀取取代規則（A欄=尋找，B欄=取代）
    lngLastRow = wsRule.Cells(wsRule.Rows.Count, 1).End(xlUp).Row
    If lngLastRow < 2 Then
        MsgBox "請在 A 欄填入尋找文字，B 欄填入取代文字（第二列起）！", vbExclamation, "提示"
        Exit Sub
    End If

    lngCount = lngLastRow - 1
    ReDim arrFind(1 To lngCount)
    ReDim arrReplace(1 To lngCount)

    For lngRow = 2 To lngLastRow
        arrFind(lngRow - 1) = CStr(wsRule.Cells(lngRow, 1).Value)
        arrReplace(lngRow - 1) = CStr(wsRule.Cells(lngRow, 2).Value)
    Next lngRow

    '選擇目標資料夾（msoFileDialogFolderPicker = 4）
    With Application.FileDialog(4)
        .Title = "請選擇含有 Word 文件的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    '建立 Word 物件（背景執行）
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    '逐一處理 .docx 檔案
    lngFileCount = 0
    strFile = Dir(strFolder & "*.docx")

    Do While strFile <> ""
        Set wdDoc = wdApp.Documents.Open(strFolder & strFile)

        '套用所有取代規則
        For i = 1 To lngCount
            If arrFind(i) <> "" Then
                With wdDoc.Content.Find
                    .ClearFormatting
                    .Replacement.ClearFormatting
                    .Text = arrFind(i)
                    .Replacement.Text = arrReplace(i)
                    .Forward = True
                    .Wrap = 1           'wdFindContinue = 1
                    .MatchCase = False
                    .Execute Replace:=2  'wdReplaceAll = 2
                End With
            End If
        Next i

        wdDoc.Save
        wdDoc.Close False
        Set wdDoc = Nothing
        lngFileCount = lngFileCount + 1
        strFile = Dir()
    Loop

    wdApp.Quit
    Set wdApp = Nothing

    MsgBox "批次取代完成！" & vbCrLf & _
           "共處理 " & lngFileCount & " 個 Word 文件。", _
           vbInformation, "完成"

    Set wsRule = Nothing
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
