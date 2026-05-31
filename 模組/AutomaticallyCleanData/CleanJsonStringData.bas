Attribute VB_Name = "CleanJsonStringData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanJsonStringData
'功能說明: 清理儲存格中的JSON格式字串，提取指定欄位值或移除JSON結構符號
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestCleanJsonStringData()
    Dim ws As Worksheet
    Set ws = GetOrCreateJsonSheet(ThisWorkbook, "JSON字串清理範例")
    Call FillJsonSampleData(ws)
    Call CleanJsonStringData(ws, 2, 1, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)
    ws.Columns.AutoFit
    MsgBox "JSON字串清理完成！", vbInformation, "完成"
End Sub

Sub CleanJsonStringData(ByVal ws As Worksheet, ByVal jsonCol As Integer, _
                         ByVal startRow As Long, ByVal endRow As Long)
    Dim i        As Long
    Dim rawStr   As String
    Dim cleanStr As String
    Dim parts()  As String
    Dim part     As Variant
    Dim result   As String
    Dim colonPos As Integer

    Application.ScreenUpdating = False

    For i = startRow To endRow
        rawStr = CStr(ws.Cells(i, jsonCol).Value)

        If Len(rawStr) > 0 Then
            cleanStr = rawStr
            cleanStr = Replace(cleanStr, "{", "")
            cleanStr = Replace(cleanStr, "}", "")
            cleanStr = Replace(cleanStr, "[", "")
            cleanStr = Replace(cleanStr, "]", "")
            cleanStr = Replace(cleanStr, Chr(34), "")
            cleanStr = Replace(cleanStr, ",", "; ")

            result = ""
            parts = Split(cleanStr, "; ")
            For Each part In parts
                colonPos = InStr(CStr(part), ":")
                If colonPos > 0 Then
                    result = result & Trim(Mid(CStr(part), colonPos + 1)) & " | "
                Else
                    result = result & Trim(CStr(part)) & " | "
                End If
            Next part

            If Right(Trim(result), 3) = " | " Then
                result = Left(Trim(result), Len(Trim(result)) - 3)
            End If

            ws.Cells(i, jsonCol + 1).Value = Trim(result)
        End If
    Next i

    Application.ScreenUpdating = True
End Sub

Private Sub FillJsonSampleData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("原始JSON", "清理後")
    ws.Range("A2").Value = "{name:Wang,dept:Sales,salary:50000}"
    ws.Range("A3").Value = "{name:Lee,dept:Marketing,salary:45000}"
    ws.Range("A4").Value = "{name:Chen,dept:HR,salary:48000}"
    ws.Range("A5").Value = "[{id:1,status:active}]"
End Sub

Private Function GetOrCreateJsonSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateJsonSheet = ws
End Function
