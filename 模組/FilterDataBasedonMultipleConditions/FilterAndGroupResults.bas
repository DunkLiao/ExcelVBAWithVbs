Option Explicit
Attribute VB_Name = "FilterAndGroupResults"
'*************************************************************************************
'模組名稱: FilterAndGroupResults
'功能說明: 依多重條件篩選資料後，自動依類別分組並統計各組小計
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/22
'
'*************************************************************************************

Sub TestFilterAndGroupResults()
    Call CreateFilterGroupDemo(ThisWorkbook)
End Sub

Sub CreateFilterGroupDemo(ByVal wb As Workbook)
    On Error GoTo ErrorHandler

    Dim srcWs As Worksheet
    Set srcWs = GetOrCreateFGSheet(wb, "銷售原始資料")
    srcWs.Cells.Clear
    Call FillGroupSalesData(srcWs)

    Dim resultWs As Worksheet
    Set resultWs = GetOrCreateFGSheet(wb, "篩選分組結果")
    resultWs.Cells.Clear

    Call FilterAndGroup(srcWs, resultWs, "業務部", 40000)

    resultWs.Columns.AutoFit
    MsgBox "篩選並分組完成，請查看篩選分組結果工作表！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "篩選分組時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

Private Sub FillGroupSalesData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("姓名", "部門", "業績", "達標")
    ws.Range("A2:D2").Value = Array("張大明", "業務部", 85000, "是")
    ws.Range("A3:D3").Value = Array("李小華", "人事部", 32000, "否")
    ws.Range("A4:D4").Value = Array("王美麗", "業務部", 62000, "是")
    ws.Range("A5:D5").Value = Array("陳志偉", "資訊部", 45000, "是")
    ws.Range("A6:D6").Value = Array("林怡君", "業務部", 38000, "否")
    ws.Range("A7:D7").Value = Array("吳建國", "資訊部", 71000, "是")
    ws.Range("A8:D8").Value = Array("黃淑芬", "業務部", 55000, "是")
    ws.Columns.AutoFit
End Sub

Private Sub FilterAndGroup( _
    ByVal srcWs As Worksheet, _
    ByVal destWs As Worksheet, _
    ByVal deptFilter As String, _
    ByVal minSales As Long)

    Dim lastRow As Long
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

    destWs.Range("A1:E1").Value = Array("姓名", "部門", "業績", "達標", "分組")
    destWs.Range("A1:E1").Font.Bold = True

    Dim destRow As Long
    destRow = 2

    Dim totalSales As Long
    totalSales = 0

    Dim groupCount As Long
    groupCount = 0

    Dim r As Long
    For r = 2 To lastRow
        Dim dept As String
        dept = CStr(srcWs.Cells(r, 2).Value)

        Dim sales As Long
        sales = CLng(srcWs.Cells(r, 3).Value)

        If dept = deptFilter And sales >= minSales Then
            destWs.Cells(destRow, 1).Value = srcWs.Cells(r, 1).Value
            destWs.Cells(destRow, 2).Value = dept
            destWs.Cells(destRow, 3).Value = sales
            destWs.Cells(destRow, 4).Value = srcWs.Cells(r, 4).Value
            destWs.Cells(destRow, 5).Value = deptFilter & "-達標"
            destWs.Cells(destRow, 3).NumberFormat = "#,##0"
            totalSales = totalSales + sales
            groupCount = groupCount + 1
            destRow = destRow + 1
        End If
    Next r

    If groupCount > 0 Then
        destWs.Cells(destRow, 1).Value = "小計"
        destWs.Cells(destRow, 3).Value = totalSales
        destWs.Cells(destRow, 3).NumberFormat = "#,##0"
        destWs.Cells(destRow, 5).Value = "共 " & groupCount & " 筆"
        destWs.Cells(destRow, 1).Resize(1, 5).Font.Bold = True
        destWs.Cells(destRow, 1).Resize(1, 5).Interior.Color = RGB(255, 242, 204)
    End If
End Sub

Private Function GetOrCreateFGSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateFGSheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateFGSheet Is Nothing Then
        Set GetOrCreateFGSheet = wb.Worksheets.Add
        GetOrCreateFGSheet.Name = sheetName
    End If
End Function
