Option Explicit
'*************************************************************************************
'模組名稱: SplitByNamedRange
'功能說明: 依照活頁簿中定義的具名範圍，將每個具名範圍的資料分割到個別工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub SplitByNamedRange()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim nm As Name
    Dim rng As Range
    Dim newWb As Workbook
    Dim nameCount As Long
    Dim sheetName As String

    On Error GoTo ErrHandler

    Set wb = ThisWorkbook
    Set ws = ActiveSheet
    nameCount = 0

    ' 檢查是否有具名範圍
    If wb.Names.Count = 0 Then
        MsgBox "活頁簿中沒有定義任何具名範圍！" & vbNewLine & _
               "請先在【公式】→【名稱管理員】中建立具名範圍後再執行。", _
               vbExclamation, "提示"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 建立新活頁簿
    Set newWb = Workbooks.Add
    Do While newWb.Sheets.Count > 1
        newWb.Sheets(newWb.Sheets.Count).Delete
    Loop

    ' 遍歷所有具名範圍
    For Each nm In wb.Names
        On Error Resume Next
        Set rng = nm.RefersToRange
        On Error GoTo ErrHandler

        If Not rng Is Nothing Then
            ' 只處理位於作用中工作表的具名範圍
            If rng.Worksheet.Name = ws.Name Then
                nameCount = nameCount + 1

                ' 清除名稱中的非法字元
                sheetName = nm.Name
                sheetName = Replace(sheetName, "!", "_")
                sheetName = Replace(sheetName, ":", "_")
                sheetName = Replace(sheetName, "\", "_")
                sheetName = Replace(sheetName, "/", "_")
                sheetName = Replace(sheetName, "?", "_")
                sheetName = Replace(sheetName, "*", "_")
                sheetName = Replace(sheetName, "[", "_")
                sheetName = Replace(sheetName, "]", "_")
                If Len(sheetName) > 31 Then sheetName = Left(sheetName, 31)

                ' 建立或取用工作表
                If nameCount = 1 Then
                    Set newWs = newWb.Sheets(1)
                    newWs.Name = sheetName
                Else
                    Set newWs = newWb.Sheets.Add(After:=newWb.Sheets(newWb.Sheets.Count))
                    newWs.Name = sheetName
                End If

                ' 複製具名範圍資料
                rng.Copy Destination:=newWs.Range("A1")

                ' 加入說明資訊
                Dim infoRow As Long
                infoRow = rng.Rows.Count + 2
                newWs.Cells(infoRow, 1).Value = "來源具名範圍: " & nm.Name
                newWs.Cells(infoRow + 1, 1).Value = "來源位址: " & rng.Address
                newWs.Cells(infoRow, 1).Font.Italic = True
                newWs.Cells(infoRow + 1, 1).Font.Italic = True

                newWs.Columns.AutoFit
            End If
        End If
        Set rng = Nothing
    Next nm

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    If nameCount = 0 Then
        newWb.Close SaveChanges:=False
        MsgBox "作用中工作表中沒有找到具名範圍！", vbExclamation, "提示"
    Else
        MsgBox "依具名範圍分割完成！共分割 " & nameCount & " 個具名範圍。", vbInformation, "完成"
    End If
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
