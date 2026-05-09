Attribute VB_Name = "SetWordPageLayout"
Option Explicit
'*************************************************************************************
'模組名稱: SetWordPageLayout
'功能說明: 批次設定資料夾內所有 Word 文件的頁面版面配置
'          （紙張大小、方向、邊界）
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期: 2026/05/10
'
'使用方式:
'  1. 執行 BatchSetWordPageLayout
'  2. 選擇紙張大小（A4 / A3 / Letter）
'  3. 選擇方向（直向 / 橫向）
'  4. 輸入上下左右邊界（公分）
'  5. 選擇目標資料夾，程式自動套用
'
'注意事項:
'  - 1 公分 = 567 twips（Word 內部單位）
'  - 預設值以 A4 直向、2.54cm 邊界為基準
'*************************************************************************************

'批次設定 Word 頁面版面
Sub BatchSetWordPageLayout()
    Const CM_TO_TWIPS As Double = 567

    Dim wdApp        As Object
    Dim wdDoc        As Object
    Dim wdSec        As Object
    Dim strFolder    As String
    Dim strFile      As String
    Dim lngPaperSize As Long
    Dim lngOrient    As Long
    Dim dblTop       As Double
    Dim dblBottom    As Double
    Dim dblLeft      As Double
    Dim dblRight     As Double
    Dim lngCount     As Long
    Dim strChoice    As String

    On Error GoTo ErrHandler

    '選擇紙張大小
    strChoice = InputBox("請輸入紙張大小：" & vbCrLf & _
                         "1 = A4" & vbCrLf & _
                         "2 = A3" & vbCrLf & _
                         "3 = Letter", "紙張大小", "1")
    Select Case strChoice
        Case "1": lngPaperSize = 9    'wdPaperA4 = 9
        Case "2": lngPaperSize = 8    'wdPaperA3 = 8
        Case "3": lngPaperSize = 1    'wdPaperLetter = 1
        Case Else
            MsgBox "無效選擇，操作取消。", vbExclamation, "提示"
            Exit Sub
    End Select

    '選擇方向
    strChoice = InputBox("請輸入頁面方向：" & vbCrLf & _
                         "0 = 直向" & vbCrLf & _
                         "1 = 橫向", "頁面方向", "0")
    Select Case strChoice
        Case "0": lngOrient = 0   'wdOrientPortrait = 0
        Case "1": lngOrient = 1   'wdOrientLandscape = 1
        Case Else
            MsgBox "無效選擇，操作取消。", vbExclamation, "提示"
            Exit Sub
    End Select

    '輸入邊界（公分）
    dblTop = CDbl(InputBox("上邊界（公分）：", "邊界設定", "2.54"))
    dblBottom = CDbl(InputBox("下邊界（公分）：", "邊界設定", "2.54"))
    dblLeft = CDbl(InputBox("左邊界（公分）：", "邊界設定", "3.17"))
    dblRight = CDbl(InputBox("右邊界（公分）：", "邊界設定", "3.17"))

    '選擇目標資料夾
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
            With wdSec.PageSetup
                .PaperSize = lngPaperSize
                .Orientation = lngOrient
                .TopMargin = dblTop * CM_TO_TWIPS
                .BottomMargin = dblBottom * CM_TO_TWIPS
                .LeftMargin = dblLeft * CM_TO_TWIPS
                .RightMargin = dblRight * CM_TO_TWIPS
            End With
        Next wdSec

        wdDoc.Save
        wdDoc.Close False
        Set wdDoc = Nothing
        lngCount = lngCount + 1
        strFile = Dir()
    Loop

    wdApp.Quit
    Set wdApp = Nothing

    MsgBox "版面設定完成！" & vbCrLf & _
           "共處理 " & lngCount & " 個文件。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
    If Not wdDoc Is Nothing Then wdDoc.Close False
    If Not wdApp Is Nothing Then wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub
