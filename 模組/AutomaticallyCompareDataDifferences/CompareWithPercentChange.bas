Option Explicit
Attribute VB_Name = "CompareWithPercentChange"
'*************************************************************************************

'模組名稱: CompareWithPercentChange

'功能說明: 比較兩欄數值資料，計算百分比變化，並以色彩標示上升、下降或持平

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub CompareWithPercentChange()

    Dim ws As Worksheet

    Dim wsResult As Worksheet

    Dim lastRow As Long

    Dim i As Long

    Dim oldVal As Double

    Dim newVal As Double

    Dim pctChange As Double

    Dim col1 As Integer

    Dim col2 As Integer

    Dim colInput As String

    Dim dstRow As Long



    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row



    If lastRow < 2 Then

        MsgBox "工作表資料不足。", vbExclamation, "提示"

        Exit Sub

    End If



    colInput = InputBox("請輸入舊值欄號（例如：1）：", "舊值欄", "1")

    If colInput = "" Then Exit Sub

    col1 = CInt(colInput)



    colInput = InputBox("請輸入新值欄號（例如：2）：", "新值欄", "2")

    If colInput = "" Then Exit Sub

    col2 = CInt(colInput)



    ' 建立結果工作表

    On Error Resume Next

    Application.DisplayAlerts = False

    ThisWorkbook.Worksheets("PercentChanges").Delete

    Application.DisplayAlerts = True

    On Error GoTo 0



    Set wsResult = ThisWorkbook.Worksheets.Add( _

        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

    wsResult.Name = "PercentChanges"



    ' 輸出標題

    wsResult.Cells(1, 1).Value = "列號"

    wsResult.Cells(1, 2).Value = ws.Cells(1, col1).Value & "（舊值）"

    wsResult.Cells(1, 3).Value = ws.Cells(1, col2).Value & "（新值）"

    wsResult.Cells(1, 4).Value = "變化量"

    wsResult.Cells(1, 5).Value = "百分比變化"

    wsResult.Cells(1, 6).Value = "趨勢"

    wsResult.Rows(1).Font.Bold = True



    Application.ScreenUpdating = False



    dstRow = 2



    For i = 2 To lastRow

        On Error Resume Next

        oldVal = CDbl(ws.Cells(i, col1).Value)

        newVal = CDbl(ws.Cells(i, col2).Value)

        On Error GoTo 0



        wsResult.Cells(dstRow, 1).Value = i

        wsResult.Cells(dstRow, 2).Value = oldVal

        wsResult.Cells(dstRow, 3).Value = newVal

        wsResult.Cells(dstRow, 4).Value = newVal - oldVal



        If oldVal <> 0 Then

            pctChange = (newVal - oldVal) / Abs(oldVal) * 100

            wsResult.Cells(dstRow, 5).Value = pctChange / 100

            wsResult.Cells(dstRow, 5).NumberFormat = "0.00%"

        Else

            wsResult.Cells(dstRow, 5).Value = "N/A"

            pctChange = 0

        End If



        ' 標示趨勢

        If newVal > oldVal Then

            wsResult.Cells(dstRow, 6).Value = "上升"

            wsResult.Rows(dstRow).Interior.Color = RGB(198, 239, 206)

            wsResult.Rows(dstRow).Font.Color = RGB(0, 97, 0)

        ElseIf newVal < oldVal Then

            wsResult.Cells(dstRow, 6).Value = "下降"

            wsResult.Rows(dstRow).Interior.Color = RGB(255, 199, 206)

            wsResult.Rows(dstRow).Font.Color = RGB(156, 0, 6)

        Else

            wsResult.Cells(dstRow, 6).Value = "持平"

            wsResult.Rows(dstRow).Interior.Color = RGB(255, 255, 153)

            wsResult.Rows(dstRow).Font.Color = RGB(102, 102, 0)

        End If



        dstRow = dstRow + 1

    Next i



    wsResult.Columns.AutoFit

    Application.ScreenUpdating = True



    MsgBox "百分比變化比較完成！共 " & (dstRow - 2) & " 筆，結果在工作表：" & wsResult.Name, _

        vbInformation, "完成"

End Sub

