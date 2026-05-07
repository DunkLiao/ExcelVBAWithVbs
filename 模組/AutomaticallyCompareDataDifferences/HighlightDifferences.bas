Attribute VB_Name = "HighlightDifferences"
Option Explicit

' ============================================================
' 範例：比較兩個工作表的相同位置資料，並以顏色標示差異儲存格
' 功能：在第二個工作表中，將與第一個工作表不同的儲存格標為黃色
' ============================================================
Sub HighlightSheetDifferences()
    Dim ws1         As Worksheet
    Dim ws2         As Worksheet
    Dim strName1    As String
    Dim strName2    As String
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim lngMaxRow   As Long
    Dim lngMaxCol   As Long
    Dim intDiffCnt  As Integer

    On Error GoTo ErrHandler

    ' --- 詢問工作表名稱 ---
    strName1 = InputBox("請輸入第一個工作表名稱（基準）：", "比較工作表", "Sheet1")
    If strName1 = "" Then Exit Sub

    strName2 = InputBox("請輸入第二個工作表名稱（比較對象）：", "比較工作表", "Sheet2")
    If strName2 = "" Then Exit Sub

    Set ws1 = ThisWorkbook.Worksheets(strName1)
    Set ws2 = ThisWorkbook.Worksheets(strName2)

    ' --- 決定比較範圍 ---
    lngMaxRow = Application.WorksheetFunction.Max( _
        ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row, _
        ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row)
    lngMaxCol = Application.WorksheetFunction.Max( _
        ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column, _
        ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column)

    ' --- 清除第二個工作表既有標色 ---
    ws2.Cells.Interior.ColorIndex = xlNone

    Application.ScreenUpdating = False
    intDiffCnt = 0

    ' --- 逐格比較並標色 ---
    For lngRow = 1 To lngMaxRow
        For lngCol = 1 To lngMaxCol
            If CStr(ws1.Cells(lngRow, lngCol).Value) <> _
               CStr(ws2.Cells(lngRow, lngCol).Value) Then
                ws2.Cells(lngRow, lngCol).Interior.Color = RGB(255, 255, 0)
                intDiffCnt = intDiffCnt + 1
            End If
        Next lngCol
    Next lngRow

    Application.ScreenUpdating = True
    MsgBox "比較完成，共發現 " & intDiffCnt & " 個差異儲存格（已在「" & _
        strName2 & "」工作表中以黃色標示）。", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
