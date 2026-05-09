Attribute VB_Name = "SplitSheetByColorTag"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetByColorTag
'功能說明: 依指定欄位的儲存格背景色分割工作表，每種顏色存成一個 Excel 檔案
'          常用於以顏色標記的優先等級或分類資料
'
'作者版權: Dunk
'程式設計: Dunk
'最後修改: 2026/5/9
'
'*************************************************************************************

' 測試用入口：建立範例顏色標記資料後執行分割
Sub TestSplitByColor()
    Dim ws As Worksheet
    Set ws = GetOrCreateSheetColor(ThisWorkbook, "顏色標記來源")
    Call FillSampleColorData(ws)
    Call SplitSheetByInteriorColor(ws, 1)
End Sub

' 依儲存格背景色拆分工作表
' ws: 來源工作表  colorColIndex: 判斷顏色的欄號
Sub SplitSheetByInteriorColor(ByVal ws As Worksheet, ByVal colorColIndex As Integer)
    Dim folderPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim colCount As Long
    Dim cellColor As Long
    Dim cellColor2 As Long
    Dim colorKey As String
    Dim uniqueColors As Collection
    Dim alreadyAdded As Boolean
    Dim key As Variant
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    Dim newRow As Long
    Dim copyRow As Long

    lastRow = ws.Cells(ws.Rows.Count, colorColIndex).End(xlUp).Row
    If lastRow <= 1 Then
        MsgBox "工作表無資料可分割！", vbExclamation, "警告"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇輸出資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作", vbInformation, "取消"
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    colCount = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    Set uniqueColors = New Collection
    For i = 2 To lastRow
        cellColor = ws.Cells(i, colorColIndex).Interior.Color
        colorKey = CStr(cellColor)
        alreadyAdded = False
        For Each key In uniqueColors
            If key = colorKey Then
                alreadyAdded = True
                Exit For
            End If
        Next key
        If Not alreadyAdded Then uniqueColors.Add colorKey
    Next i

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each key In uniqueColors
        Set wbNew = Workbooks.Add
        Set wsNew = wbNew.Worksheets(1)
        wsNew.Name = Left("Color_" & Hex(CLng(CStr(key))), 31)
        ws.Range(ws.Cells(1, 1), ws.Cells(1, colCount)).Copy _
            Destination:=wsNew.Cells(1, 1)
        newRow = 2
        For copyRow = 2 To lastRow
            cellColor2 = ws.Cells(copyRow, colorColIndex).Interior.Color
            If CStr(cellColor2) = CStr(key) Then
                ws.Range(ws.Cells(copyRow, 1), ws.Cells(copyRow, colCount)).Copy _
                    Destination:=wsNew.Cells(newRow, 1)
                newRow = newRow + 1
            End If
        Next copyRow
        wsNew.Columns.AutoFit
        wbNew.SaveAs folderPath & "Color_" & Hex(CLng(CStr(key))) & ".xlsx", xlOpenXMLWorkbook
        wbNew.Close SaveChanges:=False
    Next key

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "顏色分割完成！共建立 " & uniqueColors.Count & " 個檔案。", vbInformation, "完成"
End Sub

' 建立範例顏色標記資料（高=紅、中=黃、低=綠）
Private Sub FillSampleColorData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1").Value = "任務"
    ws.Range("B1").Value = "負責人"
    ws.Range("C1").Value = "優先度"
    ws.Range("A2").Value = "簽約審核": ws.Range("B2").Value = "王大明": ws.Range("C2").Value = "高"
    ws.Range("A3").Value = "報告撰寫": ws.Range("B3").Value = "李小美": ws.Range("C3").Value = "低"
    ws.Range("A4").Value = "客戶拜訪": ws.Range("B4").Value = "張志偉": ws.Range("C4").Value = "高"
    ws.Range("A5").Value = "會議準備": ws.Range("B5").Value = "陳美如": ws.Range("C5").Value = "中"
    ws.Range("A6").Value = "系統測試": ws.Range("B6").Value = "林正雄": ws.Range("C6").Value = "中"
    ws.Range("A7").Value = "資料備份": ws.Range("B7").Value = "吳淑芬": ws.Range("C7").Value = "低"
    ws.Range("A2:C2").Interior.Color = RGB(255, 153, 153)
    ws.Range("A3:C3").Interior.Color = RGB(153, 255, 153)
    ws.Range("A4:C4").Interior.Color = RGB(255, 153, 153)
    ws.Range("A5:C5").Interior.Color = RGB(255, 255, 153)
    ws.Range("A6:C6").Interior.Color = RGB(255, 255, 153)
    ws.Range("A7:C7").Interior.Color = RGB(153, 255, 153)
    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateSheetColor(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetColor = ws
End Function
