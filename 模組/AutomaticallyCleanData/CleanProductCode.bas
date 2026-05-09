Attribute VB_Name = "CleanProductCode"
Option Explicit
'*************************************************************************************
'模組名稱: CleanProductCode
'功能說明: 清理產品代碼欄位，統一轉為大寫、移除非字母數字字元、補足前導零至固定長度
'著作權所有: Dunk
'撰寫日期: 2026/5/9
'*************************************************************************************

' 測試入口：建立範例工作表並執行清理
Sub TestCleanProductCode()
    Dim ws As Worksheet
    Set ws = GetOrCreateProdCodeSheet(ThisWorkbook, "產品代碼清理範例")
    Call FillDirtyProductCodeData(ws)
    Call StandardizeProductCodeColumn(ws, 1, 8)
    ws.Columns("A:C").AutoFit
    MsgBox "產品代碼清理完成！", vbInformation, "完成"
End Sub

' 清理產品代碼：統一大寫、移除非英數字元、補足固定長度
Sub StandardizeProductCodeColumn(ByVal ws As Worksheet, ByVal colIndex As Long, ByVal codeLength As Integer)
    Dim lastRow  As Long
    Dim r        As Long
    Dim strVal   As String
    Dim cleanVal As String
    Dim i        As Integer
    Dim c        As String

    Application.ScreenUpdating = False
    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    For r = 2 To lastRow
        strVal = Trim(UCase(CStr(ws.Cells(r, colIndex).Value)))
        If strVal <> "" Then
            ' 移除非英數字元
            cleanVal = ""
            For i = 1 To Len(strVal)
                c = Mid(strVal, i, 1)
                If (c >= "A" And c <= "Z") Or (c >= "0" And c <= "9") Then
                    cleanVal = cleanVal & c
                End If
            Next i
            ' 補足前導零至指定長度（如果全為數字部分）
            If Len(cleanVal) > 0 And Len(cleanVal) < codeLength Then
                cleanVal = String(codeLength - Len(cleanVal), "0") & cleanVal
            End If
            ws.Cells(r, colIndex).Value = cleanVal
        End If
    Next r

    Application.ScreenUpdating = True
End Sub

Private Sub FillDirtyProductCodeData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("產品代碼", "品名", "狀態")
    ws.Range("A2:C2").Value = Array("prd-001", "鍵盤", "上架")
    ws.Range("A3:C3").Value = Array("PRD_002 ", "滑鼠", "上架")
    ws.Range("A4:C4").Value = Array("prd#3", "耳機", "缺貨")
    ws.Range("A5:C5").Value = Array("PRD 0004", "螢幕", "上架")
    ws.Range("A6:C6").Value = Array("prd.5", "鏡頭", "預購")
    ws.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreateProdCodeSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateProdCodeSheet = ws
End Function