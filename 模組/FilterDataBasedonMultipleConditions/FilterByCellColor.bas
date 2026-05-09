Attribute VB_Name = "FilterByCellColor"
Option Explicit

'************************************************************************************
' 模組名稱: FilterByCellColor
' 功能說明: 依儲存格背景色篩選資料（AutoFilter FilterByColor）
'           同時示範手動以 Interior.Color 逐列比對並複製到新工作表
'
' 作者版權: Dunk
' 現任設計: Dunk
' 最後修改: 2026/5/9
'************************************************************************************

' 定義標記色彩常數
Private Const COLOR_RED    As Long = RGB(255, 199, 206)  ' 紅色警示（庫存不足）
Private Const COLOR_YELLOW As Long = RGB(255, 235, 156)  ' 黃色提醒（即將不足）
Private Const COLOR_GREEN  As Long = RGB(198, 239, 206)  ' 綠色正常

' 入口：以 AutoFilter FilterByColor 篩選紅色警示列
Public Sub FilterByRedColorExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateWsColor(ThisWorkbook, "色彩篩選範例")
    Call FillColorCodedData(ws)

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' AutoFilter 依背景色篩選（xlFilterCellColor：背景色）
    ws.Range("A1").CurrentRegion.AutoFilter _
        Field:=4, _
        Criteria1:=COLOR_RED, _
        Operator:=xlFilterCellColor

    ws.Columns("A:D").AutoFit
    MsgBox "已篩選庫存不足（紅色警示）的商品。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 入口：手動逐列比對顏色，將紅色列複製到新工作表
Public Sub CopyRedRowsToNewSheet()
    On Error GoTo ErrHandler

    Dim wsSource As Worksheet
    Dim wsResult As Worksheet
    Dim lastRow  As Long
    Dim i        As Long
    Dim resultRow As Long

    Set wsSource = GetOrCreateWsColor(ThisWorkbook, "色彩篩選範例")
    Call FillColorCodedData(wsSource)

    Set wsResult = GetOrCreateWsColor(ThisWorkbook, "紅色警示清單")
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    ' 複製標題列
    wsSource.Rows(1).Copy Destination:=wsResult.Rows(1)
    resultRow = 2

    For i = 2 To lastRow
        ' 判斷 D 欄（庫存量）儲存格的背景色
        If wsSource.Cells(i, 4).Interior.Color = COLOR_RED Then
            wsSource.Rows(i).Copy Destination:=wsResult.Rows(resultRow)
            resultRow = resultRow + 1
        End If
    Next i

    wsResult.Columns.AutoFit
    wsResult.Activate
    MsgBox "已將紅色警示列複製到「紅色警示清單」工作表。共 " & _
           (resultRow - 2) & " 筆。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 清除色彩篩選
Public Sub ClearColorFilter()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    MsgBox "已清除色彩篩選。", vbInformation, "完成"
End Sub

' 填入帶有色彩標記的庫存資料
Private Sub FillColorCodedData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("商品代碼", "品名", "安全庫存", "現有庫存")

    Dim rowData As Variant
    rowData = Array( _
        Array("G001", "電阻 10K", 100, 250), _
        Array("G002", "電容 100uF", 200, 45), _
        Array("G003", "二極體 1N4007", 300, 320), _
        Array("G004", "電感 10uH", 150, 80), _
        Array("G005", "IC NE555", 50, 10), _
        Array("G006", "LED 紅色", 500, 620), _
        Array("G007", "繼電器 5V", 80, 60), _
        Array("G008", "光耦合器", 120, 5))

    Dim i As Integer
    For i = 0 To 7
        Dim safeQty   As Long
        Dim currentQty As Long
        safeQty = CLng(rowData(i)(2))
        currentQty = CLng(rowData(i)(3))

        ws.Cells(i + 2, 1).Value = rowData(i)(0)
        ws.Cells(i + 2, 2).Value = rowData(i)(1)
        ws.Cells(i + 2, 3).Value = safeQty
        ws.Cells(i + 2, 4).Value = currentQty

        ' 依庫存量套用色彩標記
        If currentQty < safeQty * 0.3 Then
            ws.Cells(i + 2, 4).Interior.Color = COLOR_RED     ' 嚴重不足
        ElseIf currentQty < safeQty Then
            ws.Cells(i + 2, 4).Interior.Color = COLOR_YELLOW  ' 低於安全量
        Else
            ws.Cells(i + 2, 4).Interior.Color = COLOR_GREEN   ' 正常
        End If
    Next i

    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表並清空
Private Function GetOrCreateWsColor(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateWsColor = ws
End Function