Option Explicit
Attribute VB_Name = "CleanTabData"
'*************************************************************************************
'模組名稱: CleanTabData
'功能說明: 清除儲存格內容中的定位字元（Tab 鍵 Chr(9)），修復因匯入造成的格式錯亂
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestCleanTabData()
    Dim ws As Worksheet
    Set ws = GetOrCreateTabSheet(ThisWorkbook, "Tab字元清理範例")
    Call FillTabDirtyData(ws)
    Call CleanTabCharacters(ws, ws.UsedRange)
End Sub

Sub CleanTabCharacters(ByVal ws As Worksheet, ByVal targetRange As Range)
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    Dim c As Range
    Dim cleanCount As Long
    cleanCount = 0

    For Each c In targetRange.Cells
        If InStr(CStr(c.Value), Chr(9)) > 0 Then
            c.Value = Replace(CStr(c.Value), Chr(9), " ")
            c.Value = Trim(c.Value)
            cleanCount = cleanCount + 1
        End If
    Next c

    Application.ScreenUpdating = True
    MsgBox "已清理 " & cleanCount & " 個含 Tab 字元的儲存格！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清理 Tab 字元時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillTabDirtyData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("編號", "姓名", "備註")
    ws.Range("A2").Value = "001"
    ws.Range("B2").Value = "張" & Chr(9) & "大明"
    ws.Range("C2").Value = "正常" & Chr(9) & "員工"
    ws.Range("A3").Value = "002"
    ws.Range("B3").Value = "李" & Chr(9) & Chr(9) & "小華"
    ws.Range("C3").Value = "業務" & Chr(9) & "主管"
    ws.Range("A4").Value = "003"
    ws.Range("B4").Value = "王美麗"
    ws.Range("C4").Value = "資訊" & Chr(9) & "部門"
    ws.Columns.AutoFit
End Sub

Private Function GetOrCreateTabSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateTabSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateTabSheet Is Nothing Then
        Set GetOrCreateTabSheet = wb.Worksheets.Add
        GetOrCreateTabSheet.Name = sheetName
    End If
End Function
