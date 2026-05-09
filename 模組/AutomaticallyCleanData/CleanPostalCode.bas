Attribute VB_Name = "CleanPostalCode"
Option Explicit
'*************************************************************************************
'模組名稱: CleanPostalCode
'功能說明: 清理郵遞區號欄位，移除非數字字元，並補足前導零至 3 位（臺灣郵遞區號格式）
'著作權所有: Dunk
'撰寫日期: 2026/5/9
'*************************************************************************************

Sub TestCleanPostalCode()
    Dim ws As Worksheet
    Set ws = GetOrCreatePostalSheet(ThisWorkbook, "郵遞區號清理範例")
    Call FillDirtyPostalData(ws)
    Call CleanPostalCodeColumn(ws, 3, 3)
    ws.Columns("A:C").AutoFit
    MsgBox "郵遞區號清理完成！", vbInformation, "完成"
End Sub

' 清理郵遞區號欄位，移除非數字並補足指定位數
Sub CleanPostalCodeColumn(ByVal ws As Worksheet, ByVal colIndex As Long, ByVal digitCount As Integer)
    Dim lastRow  As Long
    Dim r        As Long
    Dim strVal   As String
    Dim cleanVal As String
    Dim i        As Integer
    Dim c        As String

    Application.ScreenUpdating = False
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        strVal = Trim(CStr(ws.Cells(r, colIndex).Value))
        If strVal <> "" Then
            ' 移除非數字字元
            cleanVal = ""
            For i = 1 To Len(strVal)
                c = Mid(strVal, i, 1)
                If c >= "0" And c <= "9" Then
                    cleanVal = cleanVal & c
                End If
            Next i
            ' 補足前導零至指定位數
            If Len(cleanVal) > 0 And Len(cleanVal) < digitCount Then
                cleanVal = String(digitCount - Len(cleanVal), "0") & cleanVal
            End If
            If Len(cleanVal) > digitCount Then
                ' 截取前 digitCount 位
                cleanVal = Left(cleanVal, digitCount)
            End If
            ws.Cells(r, colIndex).Value = cleanVal
            ws.Cells(r, colIndex).NumberFormat = "@"
        End If
    Next r

    Application.ScreenUpdating = True
End Sub

Private Sub FillDirtyPostalData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("姓名", "地址", "郵遞區號")
    ws.Range("A2:C2").Value = Array("王大明", "臺北市信義區忠孝東路五段1號", "110")
    ws.Range("A3:C3").Value = Array("李小花", "臺中市西屯區文心路100號", "40756")
    ws.Range("A4:C4").Value = Array("陳志強", "高雄市三民區九如二路200號", " 807 ")
    ws.Range("A5:C5").Value = Array("林美雪", "臺南市東區中華東路二段50號", "70")
    ws.Range("A6:C6").Value = Array("張建國", "新北市板橋區中山路一段100號", "22041-ABC")
    ws.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreatePostalSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreatePostalSheet = ws
End Function