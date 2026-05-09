Attribute VB_Name = "CleanAddressData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanAddressData
'功能說明: 清理地址欄位，補全縣市簡稱（台→臺）、移除多餘空白與全形空格
'著作權所有: Dunk
'撰寫日期: 2026/5/9
'*************************************************************************************

Sub TestCleanAddressData()
    Dim ws As Worksheet
    Set ws = GetOrCreateAddressSheet(ThisWorkbook, "地址清理範例")
    Call FillDirtyAddressData(ws)
    Call CleanAddressColumn(ws, 2)
    MsgBox "地址欄位清理完成！", vbInformation, "完成"
End Sub

' 清理地址欄位：補全縣市名稱、移除多餘空白
Sub CleanAddressColumn(ByVal ws As Worksheet, ByVal colIndex As Long)
    Dim lastRow As Long
    Dim r       As Long
    Dim strVal  As String

    Application.ScreenUpdating = False
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        strVal = Trim(CStr(ws.Cells(r, colIndex).Value))
        If strVal <> "" Then
            ' 移除全形空格（Unicode 12288）
            strVal = Replace(strVal, Chr(12288), "")
            ' 統一「台」開頭縣市名稱為「臺」
            strVal = Replace(strVal, "台北市", "臺北市")
            strVal = Replace(strVal, "台中市", "臺中市")
            strVal = Replace(strVal, "台南市", "臺南市")
            strVal = Replace(strVal, "台東縣", "臺東縣")
            ' 移除地址內多餘空白
            Do While InStr(strVal, "  ") > 0
                strVal = Replace(strVal, "  ", " ")
            Loop
            strVal = Trim(strVal)
            ws.Cells(r, colIndex).Value = strVal
        End If
    Next r

    Application.ScreenUpdating = True
End Sub

Private Sub FillDirtyAddressData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("姓名", "地址")
    ws.Range("A2:B2").Value = Array("王大明", "  台北市信義區  忠孝東路五段1號  ")
    ws.Range("A3:B3").Value = Array("李小花", "台中市西屯區文心路100號")
    ws.Range("A4:B4").Value = Array("陳志強", "台南市東區中華東路二段50號3F")
    ws.Range("A5:B5").Value = Array("林美雪", "高雄市三民區  九如二路200號")
    ws.Range("A6:B6").Value = Array("張建國", "台東縣台東市中正路1段99號")
    ws.Columns("A:B").AutoFit
End Sub

Private Function GetOrCreateAddressSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateAddressSheet = ws
End Function