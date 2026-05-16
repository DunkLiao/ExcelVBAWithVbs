Attribute VB_Name = "CompareAndColorCodeResult"
Option Explicit
'*************************************************************************************
'模組名稱: CompareAndColorCodeResult
'功能說明: 比較兩欄（或兩個範圍）的對應儲存格數值，
'          以顏色標記比較結果：
'          相同 → 綠底、增加 → 藍底、減少 → 紅底、空白差異 → 黃底
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

' 主程式：比較指定兩欄並以顏色標記差異
Sub CompareColumnsWithColorCode()
    Dim ws          As Worksheet
    Dim col1        As Long
    Dim col2        As Long
    Dim lastRow     As Long
    Dim r           As Long
    Dim val1        As Variant
    Dim val2        As Variant

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "沒有足夠的資料列。", vbExclamation
        Exit Sub
    End If

    ' 輸入兩個欄位代號
    Dim input1 As String
    Dim input2 As String
    input1 = InputBox("請輸入第一欄的欄位代號（舊資料，例如 B）：", "比較設定", "B")
    If input1 = "" Then Exit Sub
    input2 = InputBox("請輸入第二欄的欄位代號（新資料，例如 C）：", "比較設定", "C")
    If input2 = "" Then Exit Sub

    col1 = ws.Range(input1 & "1").Column
    col2 = ws.Range(input2 & "1").Column

    ' 清除舊背景色
    ws.Range(ws.Cells(2, col1), ws.Cells(lastRow, col1)).Interior.ColorIndex = xlNone
    ws.Range(ws.Cells(2, col2), ws.Cells(lastRow, col2)).Interior.ColorIndex = xlNone

    Dim cntSame    As Long
    Dim cntIncrease As Long
    Dim cntDecrease As Long
    Dim cntBlank   As Long
    cntSame = 0: cntIncrease = 0: cntDecrease = 0: cntBlank = 0

    Application.ScreenUpdating = False

    For r = 2 To lastRow
        val1 = ws.Cells(r, col1).Value
        val2 = ws.Cells(r, col2).Value

        If val1 = "" Or val2 = "" Then
            ' 有空白 → 黃底
            ws.Cells(r, col1).Interior.Color = RGB(255, 255, 0)
            ws.Cells(r, col2).Interior.Color = RGB(255, 255, 0)
            cntBlank = cntBlank + 1
        ElseIf IsNumeric(val1) And IsNumeric(val2) Then
            Dim n1 As Double
            Dim n2 As Double
            n1 = CDbl(val1)
            n2 = CDbl(val2)
            If n1 = n2 Then
                ' 相同 → 綠底
                ws.Cells(r, col1).Interior.Color = RGB(0, 176, 80)
                ws.Cells(r, col2).Interior.Color = RGB(0, 176, 80)
                cntSame = cntSame + 1
            ElseIf n2 > n1 Then
                ' 增加 → 藍底
                ws.Cells(r, col1).Interior.Color = RGB(68, 114, 196)
                ws.Cells(r, col2).Interior.Color = RGB(68, 114, 196)
                ws.Cells(r, col2).Font.Color = RGB(255, 255, 255)
                cntIncrease = cntIncrease + 1
            Else
                ' 減少 → 紅底
                ws.Cells(r, col1).Interior.Color = RGB(192, 0, 0)
                ws.Cells(r, col2).Interior.Color = RGB(192, 0, 0)
                ws.Cells(r, col1).Font.Color = RGB(255, 255, 255)
                ws.Cells(r, col2).Font.Color = RGB(255, 255, 255)
                cntDecrease = cntDecrease + 1
            End If
        Else
            ' 文字比較
            If CStr(val1) = CStr(val2) Then
                ws.Cells(r, col1).Interior.Color = RGB(0, 176, 80)
                ws.Cells(r, col2).Interior.Color = RGB(0, 176, 80)
                cntSame = cntSame + 1
            Else
                ws.Cells(r, col1).Interior.Color = RGB(192, 0, 0)
                ws.Cells(r, col2).Interior.Color = RGB(192, 0, 0)
                ws.Cells(r, col1).Font.Color = RGB(255, 255, 255)
                ws.Cells(r, col2).Font.Color = RGB(255, 255, 255)
                cntDecrease = cntDecrease + 1
            End If
        End If
    Next r

    Application.ScreenUpdating = True

    MsgBox "比較完成！" & vbCrLf & _
           "相同（綠）：" & cntSame & " 筆" & vbCrLf & _
           "增加（藍）：" & cntIncrease & " 筆" & vbCrLf & _
           "減少（紅）：" & cntDecrease & " 筆" & vbCrLf & _
           "空白（黃）：" & cntBlank & " 筆", vbInformation, "比較結果"
End Sub

' 建立示範資料並執行比較
Sub DemoCompareAndColorCode()
    Dim ws As Worksheet
    Set ws = GetOrCreateColorSheet("顏色比較示範")
    ws.Cells.Clear

    ws.Range("A1:C1").Value = Array("項目", "舊數值", "新數值")
    ws.Range("A1:C1").Font.Bold = True

    ws.Range("A2:C2").Value = Array("項目A", 1000, 1000)
    ws.Range("A3:C3").Value = Array("項目B", 850, 920)
    ws.Range("A4:C4").Value = Array("項目C", 1200, 1050)
    ws.Range("A5:C5").Value = Array("項目D", 760, 760)
    ws.Range("A6:C6").Value = Array("項目E", 430, 680)
    ws.Range("A7:C7").Value = Array("項目F", 990, "")

    ws.Columns("A:C").AutoFit

    ' 直接執行顏色比較（B vs C 欄，第2~7列）
    Dim lastRow As Long
    lastRow = 7
    Dim r As Long
    Dim val1 As Variant
    Dim val2 As Variant

    Application.ScreenUpdating = False
    For r = 2 To lastRow
        val1 = ws.Cells(r, 2).Value
        val2 = ws.Cells(r, 3).Value

        If CStr(val1) = "" Or CStr(val2) = "" Then
            ws.Cells(r, 2).Interior.Color = RGB(255, 255, 0)
            ws.Cells(r, 3).Interior.Color = RGB(255, 255, 0)
        ElseIf IsNumeric(val1) And IsNumeric(val2) Then
            If CDbl(val1) = CDbl(val2) Then
                ws.Cells(r, 2).Interior.Color = RGB(0, 176, 80)
                ws.Cells(r, 3).Interior.Color = RGB(0, 176, 80)
            ElseIf CDbl(val2) > CDbl(val1) Then
                ws.Cells(r, 2).Interior.Color = RGB(68, 114, 196)
                ws.Cells(r, 3).Interior.Color = RGB(68, 114, 196)
                ws.Cells(r, 3).Font.Color = RGB(255, 255, 255)
            Else
                ws.Cells(r, 2).Interior.Color = RGB(192, 0, 0)
                ws.Cells(r, 3).Interior.Color = RGB(192, 0, 0)
                ws.Cells(r, 2).Font.Color = RGB(255, 255, 255)
                ws.Cells(r, 3).Font.Color = RGB(255, 255, 255)
            End If
        End If
    Next r
    Application.ScreenUpdating = True

    MsgBox "示範完成！綠=相同 藍=增加 紅=減少 黃=空白", vbInformation, "完成"
End Sub

Private Function GetOrCreateColorSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateColorSheet = ws
End Function
