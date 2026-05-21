Option Explicit
Attribute VB_Name = "MergeWithDataTypeConvert"
'*************************************************************************************
'模組名稱: MergeWithDataTypeConvert
'功能說明: 合併多個工作表資料，同時自動偵測並轉換欄位資料型別
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestMergeWithDataTypeConvert()
    Call CreateMergeConvertDemo(ThisWorkbook)
End Sub

Sub CreateMergeConvertDemo(ByVal wb As Workbook)
    On Error GoTo ErrorHandler

    Call SetupMergeSourceSheets(wb)

    Dim destWs As Worksheet
    Set destWs = GetOrCreateMDCSheet(wb, "合併轉換結果")
    destWs.Cells.Clear

    Call MergeAndConvertData(wb, destWs)

    destWs.Columns.AutoFit
    MsgBox "合併並轉換資料型別完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "合併轉換時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub SetupMergeSourceSheets(ByVal wb As Workbook)
    Dim ws1 As Worksheet
    Set ws1 = GetOrCreateMDCSheet(wb, "來源資料A")
    ws1.Cells.Clear
    ws1.Range("A1:C1").Value = Array("姓名", "年齡", "加入日期")
    ws1.Range("A2:C2").Value = Array("張大明", "28", "2023/3/15")
    ws1.Range("A3:C3").Value = Array("李小華", "35", "2022/7/1")

    Dim ws2 As Worksheet
    Set ws2 = GetOrCreateMDCSheet(wb, "來源資料B")
    ws2.Cells.Clear
    ws2.Range("A1:C1").Value = Array("姓名", "年齡", "加入日期")
    ws2.Range("A2:C2").Value = Array("王美麗", "42", "2021/11/20")
    ws2.Range("A3:C3").Value = Array("陳志偉", "31", "2024/1/5")
End Sub

Private Sub MergeAndConvertData(ByVal wb As Workbook, ByVal destWs As Worksheet)
    Dim srcNames As Variant
    srcNames = Array("來源資料A", "來源資料B")

    Dim nextRow As Long
    nextRow = 1

    Dim i As Integer
    For i = 0 To UBound(srcNames)
        Dim srcWs As Worksheet
        On Error Resume Next
        Set srcWs = wb.Worksheets(CStr(srcNames(i)))
        On Error GoTo 0

        If Not srcWs Is Nothing Then
            Dim lastRow As Long
            lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

            Dim lastCol As Long
            lastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

            Dim r As Long
            For r = 1 To lastRow
                Dim c As Long
                For c = 1 To lastCol
                    Dim srcVal As Variant
                    srcVal = srcWs.Cells(r, c).Value
                    destWs.Cells(nextRow, c).Value = ConvertCellDataType(srcVal)
                Next c
                nextRow = nextRow + 1
            Next r

            Set srcWs = Nothing
        End If
    Next i
End Sub

Private Function ConvertCellDataType(ByVal val As Variant) As Variant
    If IsNumeric(val) Then
        ConvertCellDataType = CLng(val)
    ElseIf IsDate(val) Then
        ConvertCellDataType = CDate(val)
    Else
        ConvertCellDataType = CStr(val)
    End If
End Function

Private Function GetOrCreateMDCSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateMDCSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateMDCSheet Is Nothing Then
        Set GetOrCreateMDCSheet = wb.Worksheets.Add
        GetOrCreateMDCSheet.Name = sheetName
    End If
End Function
