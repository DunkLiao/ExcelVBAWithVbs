Attribute VB_Name = "CompareTwoWorkbooks"
Option Explicit
'*************************************************************************************
'模組名稱: CompareTwoWorkbooks
'功能說明: 讓使用者選取兩個 Excel 活頁簿，比對指定工作表的差異，
'          結果輸出至目前活頁簿的新工作表
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 選取兩個活頁簿後進行比對
Sub CompareTwoWorkbooks()
    Dim pathA     As String
    Dim pathB     As String
    Dim sheetName As String
    Dim wbA       As Workbook
    Dim wbB       As Workbook
    Dim wsA       As Worksheet
    Dim wsB       As Worksheet
    Dim wsR       As Worksheet
    Dim lastRow   As Long
    Dim lastCol   As Long
    Dim r         As Long
    Dim c         As Long
    Dim rptRow    As Long
    Dim valA      As String
    Dim valB      As String
    Dim diffCount As Long

    On Error GoTo ErrHandler

    ' 選取第一個活頁簿（基準）
    pathA = Application.GetOpenFilename( _
        "Excel 檔案 (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", _
        1, "選取第一個活頁簿（基準）", , False)
    If pathA = "False" Then Exit Sub

    ' 選取第二個活頁簿（比對對象）
    pathB = Application.GetOpenFilename( _
        "Excel 檔案 (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", _
        1, "選取第二個活頁簿（比對對象）", , False)
    If pathB = "False" Then Exit Sub

    ' 詢問要比對的工作表名稱
    sheetName = InputBox("請輸入要比對的工作表名稱：", "工作表名稱", "Sheet1")
    If sheetName = "" Then Exit Sub

    Application.ScreenUpdating = False

    Set wbA = Workbooks.Open(pathA, ReadOnly:=True)
    Set wbB = Workbooks.Open(pathB, ReadOnly:=True)

    On Error Resume Next
    Set wsA = wbA.Worksheets(sheetName)
    Set wsB = wbB.Worksheets(sheetName)
    On Error GoTo ErrHandler

    If wsA Is Nothing Or wsB Is Nothing Then
        MsgBox "找不到工作表「" & sheetName & "」，請確認名稱是否正確。", vbExclamation, "錯誤"
        GoTo Cleanup
    End If

    ' 在目前活頁簿建立差異報告
    Set wsR = GetOrCreateSheetCTW(ThisWorkbook, "跨檔差異報告")

    wsR.Range("A1").Value = "檔案A"
    wsR.Range("B1").Value = pathA
    wsR.Range("A2").Value = "檔案B"
    wsR.Range("B2").Value = pathB
    wsR.Range("A3").Value = "工作表"
    wsR.Range("B3").Value = sheetName
    wsR.Range("A5").Value = "儲存格"
    wsR.Range("B5").Value = "檔案A 值"
    wsR.Range("C5").Value = "檔案B 值"
    With wsR.Range("A5:C5")
        .Font.Bold = True
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
    End With

    lastRow = Application.WorksheetFunction.Max( _
        wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row, _
        wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row)
    lastCol = Application.WorksheetFunction.Max( _
        wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column, _
        wsB.Cells(1, wsB.Columns.Count).End(xlToLeft).Column)

    rptRow = 6
    diffCount = 0

    For r = 1 To lastRow
        For c = 1 To lastCol
            valA = CStr(wsA.Cells(r, c).Value)
            valB = CStr(wsB.Cells(r, c).Value)
            If valA <> valB Then
                wsR.Cells(rptRow, 1).Value = wsA.Cells(r, c).Address(False, False)
                wsR.Cells(rptRow, 2).Value = valA
                wsR.Cells(rptRow, 3).Value = valB
                wsR.Cells(rptRow, 1).Resize(1, 3).Interior.Color = RGB(255, 255, 153)
                rptRow = rptRow + 1
                diffCount = diffCount + 1
            End If
        Next c
    Next r

    wsR.Columns("A:C").AutoFit
    wsR.Activate

Cleanup:
    wbA.Close SaveChanges:=False
    wbB.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "跨檔比對完成！共發現 " & diffCount & " 處差異。", vbInformation, "比對結果"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤: " & Err.Description, vbCritical, "錯誤"
    On Error Resume Next
    If Not wbA Is Nothing Then wbA.Close SaveChanges:=False
    If Not wbB Is Nothing Then wbB.Close SaveChanges:=False
End Sub

Private Function GetOrCreateSheetCTW(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetCTW = ws
End Function
