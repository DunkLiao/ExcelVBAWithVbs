Option Explicit
'*************************************************************************************
'模組名稱: CompareByDataType
'功能說明: 比較兩個工作表對應儲存格的資料型別，標示型別不一致的差異位置並輸出報告
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Function GetDataTypeName(ByVal cellVal As Variant) As String
    ' 取得儲存格值的資料型別中文名稱
    Select Case VarType(cellVal)
        Case vbEmpty:    GetDataTypeName = "空值"
        Case vbNull:     GetDataTypeName = "Null"
        Case vbInteger:  GetDataTypeName = "整數"
        Case vbLong:     GetDataTypeName = "長整數"
        Case vbSingle:   GetDataTypeName = "單精度"
        Case vbDouble:   GetDataTypeName = "雙精度"
        Case vbCurrency: GetDataTypeName = "貨幣"
        Case vbDate:     GetDataTypeName = "日期"
        Case vbString:   GetDataTypeName = "文字"
        Case vbBoolean:  GetDataTypeName = "布林"
        Case vbError:    GetDataTypeName = "錯誤"
        Case Else:       GetDataTypeName = "其他(" & VarType(cellVal) & ")"
    End Select
End Function

Sub CompareByDataType()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim wsResult As Worksheet
    Dim maxRow As Long
    Dim maxCol As Long
    Dim i As Long
    Dim j As Long
    Dim diffCount As Long
    Dim type1 As String
    Dim type2 As String
    Dim resultRow As Long
    Dim shName As String

    On Error GoTo ErrHandler

    If ThisWorkbook.Sheets.Count < 2 Then
        MsgBox "需要至少兩個工作表才能進行比較！", vbExclamation, "提示"
        Exit Sub
    End If

    ' 選擇第一個工作表
    shName = InputBox("請輸入第一個工作表名稱：", "選擇工作表", ThisWorkbook.Sheets(1).Name)
    If shName = "" Then Exit Sub
    On Error Resume Next
    Set ws1 = ThisWorkbook.Sheets(shName)
    On Error GoTo ErrHandler
    If ws1 Is Nothing Then
        MsgBox "找不到工作表：" & shName, vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 選擇第二個工作表
    shName = InputBox("請輸入第二個工作表名稱：", "選擇工作表", ThisWorkbook.Sheets(2).Name)
    If shName = "" Then Exit Sub
    On Error Resume Next
    Set ws2 = ThisWorkbook.Sheets(shName)
    On Error GoTo ErrHandler
    If ws2 Is Nothing Then
        MsgBox "找不到工作表：" & shName, vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 建立比較報告工作表
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("型別差異報告").Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set wsResult = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResult.Name = "型別差異報告"

    ' 標題列
    wsResult.Cells(1, 1).Value = "列"
    wsResult.Cells(1, 2).Value = "欄"
    wsResult.Cells(1, 3).Value = "位址"
    wsResult.Cells(1, 4).Value = ws1.Name & " 型別"
    wsResult.Cells(1, 5).Value = ws2.Name & " 型別"
    wsResult.Cells(1, 6).Value = ws1.Name & " 值"
    wsResult.Cells(1, 7).Value = ws2.Name & " 值"
    wsResult.Rows(1).Font.Bold = True
    wsResult.Rows(1).Interior.Color = RGB(68, 114, 196)
    wsResult.Rows(1).Font.Color = RGB(255, 255, 255)
    resultRow = 2

    ' 比較範圍
    maxRow = Application.WorksheetFunction.Max( _
        ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row, _
        ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row)
    maxCol = Application.WorksheetFunction.Max( _
        ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column, _
        ws2.Cells(1, ws2.Columns.Count).End(xlToLeft).Column)

    Application.ScreenUpdating = False
    diffCount = 0

    For i = 1 To maxRow
        For j = 1 To maxCol
            type1 = GetDataTypeName(ws1.Cells(i, j).Value)
            type2 = GetDataTypeName(ws2.Cells(i, j).Value)

            If type1 <> type2 Then
                wsResult.Cells(resultRow, 1).Value = i
                wsResult.Cells(resultRow, 2).Value = j
                wsResult.Cells(resultRow, 3).Value = ws1.Cells(i, j).Address(False, False)
                wsResult.Cells(resultRow, 4).Value = type1
                wsResult.Cells(resultRow, 5).Value = type2
                wsResult.Cells(resultRow, 6).Value = CStr(ws1.Cells(i, j).Value)
                wsResult.Cells(resultRow, 7).Value = CStr(ws2.Cells(i, j).Value)

                wsResult.Rows(resultRow).Interior.Color = RGB(255, 255, 153)

                ' 在原始工作表中標記差異
                ws1.Cells(i, j).Interior.Color = RGB(255, 200, 200)
                ws2.Cells(i, j).Interior.Color = RGB(200, 200, 255)

                resultRow = resultRow + 1
                diffCount = diffCount + 1
            End If
        Next j
    Next i

    wsResult.Columns.AutoFit
    Application.ScreenUpdating = True

    If diffCount = 0 Then
        wsResult.Cells(2, 1).Value = "兩個工作表的所有儲存格型別完全一致！"
        MsgBox "比較完成！兩個工作表的資料型別完全一致。", vbInformation, "完成"
    Else
        MsgBox "比較完成！共發現 " & diffCount & " 個型別不一致的儲存格。" & vbNewLine & _
               "詳見「型別差異報告」工作表。", vbExclamation, "差異報告"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
