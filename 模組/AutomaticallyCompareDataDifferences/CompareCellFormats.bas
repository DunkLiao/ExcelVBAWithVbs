Attribute VB_Name = "CompareCellFormats"
Option Explicit
'*************************************************************************************
'模組名稱: CompareCellFormats
'功能說明: 比較兩個相同大小範圍的儲存格格式差異（字型、背景、數值格式、粗體）
'          並以橙色標示有差異的儲存格
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub CompareCellFormats()
    Dim ws          As Worksheet
    Dim rng1Addr    As String
    Dim rng2Addr    As String
    Dim rng1        As Range
    Dim rng2        As Range
    Dim c1          As Range
    Dim c2          As Range
    Dim diffCount   As Long
    Dim r           As Long
    Dim c           As Long

    Set ws = ActiveSheet

    rng1Addr = InputBox("請輸入第一個比較範圍（例如：A1:D5）：", "設定範圍1", "A1:D5")
    If rng1Addr = "" Then Exit Sub

    rng2Addr = InputBox("請輸入第二個比較範圍（大小需與範圍1相同）：", "設定範圍2", "F1:I5")
    If rng2Addr = "" Then Exit Sub

    On Error GoTo ErrHandler
    Set rng1 = ws.Range(rng1Addr)
    Set rng2 = ws.Range(rng2Addr)

    If rng1.Rows.Count <> rng2.Rows.Count Or rng1.Columns.Count <> rng2.Columns.Count Then
        MsgBox "兩個範圍大小不同，請重新設定。", vbExclamation, "錯誤"
        Exit Sub
    End If

    diffCount = 0
    Application.ScreenUpdating = False

    For r = 1 To rng1.Rows.Count
        For c = 1 To rng1.Columns.Count
            Set c1 = rng1.Cells(r, c)
            Set c2 = rng2.Cells(r, c)

            Dim hasDiff As Boolean
            hasDiff = False

            If c1.Font.Bold <> c2.Font.Bold Then hasDiff = True
            If c1.Font.Name <> c2.Font.Name Then hasDiff = True
            If c1.Font.Size <> c2.Font.Size Then hasDiff = True
            If c1.Interior.ColorIndex <> c2.Interior.ColorIndex Then hasDiff = True
            If c1.NumberFormat <> c2.NumberFormat Then hasDiff = True

            If hasDiff Then
                c1.Interior.Color = RGB(255, 165, 0)
                c2.Interior.Color = RGB(255, 165, 0)
                diffCount = diffCount + 1
            End If
        Next c
    Next r

    Application.ScreenUpdating = True
    MsgBox "格式比較完成，共發現 " & diffCount & " 個差異儲存格（以橙色標示）。", _
        vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub
