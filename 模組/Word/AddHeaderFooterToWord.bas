Attribute VB_Name = "AddHeaderFooterToWord"
Option Explicit
'*************************************************************************************
'模組名稱: AddHeaderFooterToWord
'功能說明: 批次為資料夾內所有 Word 文件加入頁首與頁尾
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 執行 BatchAddHeaderFooter
'  2. 輸入頁首文字（留空則略過頁首）
'  3. 輸入頁尾文字，可使用 {page}/{pages} 代表頁碼/總頁數
'  4. 選擇目標資料夾，程式自動處理所有 .docx
'
'注意事項:
'  - 將覆蓋原有頁首頁尾設定
'  - {page} 與 {pages} 為特殊標記，會替換為 Word 域
'*************************************************************************************

'批次加入頁首頁尾
Sub BatchAddHeaderFooter()
    Dim wdApp        As Object
    Dim wdDoc        As Object
    Dim wdSec        As Object
    Dim strFolder    As String
    Dim strFile      As String
    Dim strHeader    As String
    Dim strFooter    As String
    Dim lngCount     As Long

    On Error GoTo ErrHandler

    strHeader = InputBox("請輸入頁首文字（留空略過）：", "設定頁首", "機密文件")
    strFooter = InputBox("請輸入頁尾文字（{page}=頁碼, {pages}=總頁數）：", _
                         "設定頁尾", "第 {page} 頁，共 {pages} 頁")

    If strHeader = "" And strFooter = "" Then
        MsgBox "頁首與頁尾均為空白，操作取消。", vbInformation, "提示"
        Exit Sub
    End If

    With Application.FileDialog(4)
        .Title = "請選擇 Word 文件資料夾"
        If .Show <> -1 Then Exit Sub
        strFolder = .SelectedItems(1)
    End With
    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False

    lngCount = 0
    strFile = Dir(strFolder & "*.docx")

    Do While strFile <> ""
        Set wdDoc = wdApp.Documents.Open(strFolder & strFile)

        For Each wdSec In wdDoc.Sections
            '設定頁首
            If strHeader <> "" Then
                wdSec.Headers(1).LinkToPrevious = False
                wdSec.Headers(1).Range.Text = strHeader
            End If

            '設定頁尾（處理 {page} 與 {pages} 標記）
            If strFooter <> "" Then
                wdSec.Footers(1).LinkToPrevious = False

                Dim rng As Object
                Set rng = wdSec.Footers(1).Range

                '清空頁尾
                rng.Text = ""

                '依標記分段插入一般文字或域
                Dim strParts() As String
                Dim strFull    As String
                Dim k          As Long
                strFull = strFooter

                '用臨時符號替換 {page} 與 {pages}
                strFull = Replace(strFull, "{pages}", Chr(2))
                strFull = Replace(strFull, "{page}", Chr(1))
                strParts = Split(strFull, "")

                Dim partRng As Object
                Set partRng = wdSec.Footers(1).Range
                partRng.Collapse 0   'wdCollapseEnd

                Dim c As Long
                For c = 1 To Len(strFull)
                    Dim ch As String
                    ch = Mid(strFull, c, 1)
                    Select Case Asc(ch)
                        Case 1  '{page}
                            partRng.InsertAfter ""
                            partRng.Collapse 0
                            wdDoc.Fields.Add Range:=partRng, Type:=33  'wdFieldPage=33
                            partRng.Collapse 0
                        Case 2  '{pages}
                            partRng.InsertAfter ""
                            partRng.Collapse 0
                            wdDoc.Fields.Add Range:=partRng, Type:=26  'wdFieldNumPages=26
                            partRng.Collapse 0
                        Case Else
                            partRng.InsertAfter ch
                            partRng.Collapse 0
                    End Select
                Next c
                Set partRng = Nothing
                Set rng = Nothing
            End If
        Next wdSec

        wdDoc.Save
        wdDoc.Close False
        Set wdDoc = Nothing
        lngCount = lngCount + 1
        strFile = Dir()
    Loop

    wdApp.Quit
    Set wdApp = Nothing

    MsgBox "頁首頁尾設定完成！" & vbCrLf & _
           "共處理 " & lngCount & " 個文件。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
