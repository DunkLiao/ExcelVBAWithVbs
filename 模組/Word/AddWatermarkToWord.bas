Attribute VB_Name = "AddWatermarkToWord"
Option Explicit
'*************************************************************************************
'模組名稱: AddWatermarkToWord
'功能說明: 批次為資料夾內所有 Word 文件加入文字浮水印
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 執行 BatchAddWatermark
'  2. 輸入浮水印文字（例如：機密、DRAFT）
'  3. 選擇目標資料夾，程式自動為所有 .docx 加上浮水印
'
'注意事項:
'  - 浮水印以 WordArt 文字方塊置於頁首實現
'  - 將覆蓋原有頁首中的浮水印形狀
'*************************************************************************************

'批次加入文字浮水印
Sub BatchAddWatermark()
    Dim wdApp        As Object
    Dim wdDoc        As Object
    Dim wdShape      As Object
    Dim wdHdr        As Object
    Dim strFolder    As String
    Dim strFile      As String
    Dim strWmText    As String
    Dim lngCount     As Long

    On Error GoTo ErrHandler

    strWmText = InputBox("請輸入浮水印文字：", "浮水印設定", "機密文件")
    If strWmText = "" Then
        MsgBox "浮水印文字為空，操作取消。", vbInformation, "提示"
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

        '在第一節頁首中插入浮水印
        Set wdHdr = wdDoc.Sections(1).Headers(1)
        wdHdr.LinkToPrevious = False

        '移除頁首中既有的形狀
        Dim s As Long
        For s = wdHdr.Shapes.Count To 1 Step -1
            wdHdr.Shapes(s).Delete
        Next s

        '新增文字方塊作為浮水印
        Set wdShape = wdHdr.Shapes.AddTextbox( _
            Orientation:=1, _
            Left:=100, Top:=200, Width:=350, Height:=150)

        With wdShape
            .Line.Visible = False
            .Fill.Visible = False
            .Rotation = 315
            .TextFrame.TextRange.Text = strWmText
            With .TextFrame.TextRange.Font
                .Size = 72
                .Bold = True
                .Color = RGB(192, 192, 192)
            End With
            .TextFrame.WordWrap = False
        End With

        Set wdShape = Nothing
        Set wdHdr = Nothing

        wdDoc.Save
        wdDoc.Close False
        Set wdDoc = Nothing
        lngCount = lngCount + 1
        strFile = Dir()
    Loop

    wdApp.Quit
    Set wdApp = Nothing

    MsgBox "浮水印加入完成！" & vbCrLf & _
           "共處理 " & lngCount & " 個文件。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdShape = Nothing
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
