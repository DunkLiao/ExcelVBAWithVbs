Attribute VB_Name = "FilterUniqueRecords"
Option Explicit

'************************************************************************************
' 模組名稱: FilterUniqueRecords
' 功能說明: 使用 AdvancedFilter Unique:=True 提取不重複記錄
'           可選擇提取整列唯一值或指定欄位唯一清單
'
' 作者版權: Dunk
' 現任設計: Dunk
' 最後修改: 2026/5/9
'************************************************************************************

' 入口：提取整列不重複記錄
Public Sub ExtractUniqueRowsExample()
    On Error GoTo ErrHandler

    Dim wsData   As Worksheet
    Dim wsResult As Worksheet

    Set wsData = GetOrCreateWsUnique(ThisWorkbook, "去重來源資料")
    Call FillDuplicateData(wsData)

    Set wsResult = GetOrCreateWsUnique(ThisWorkbook, "唯一記錄結果")

    ' AdvancedFilter 去重複並複製到結果工作表
    wsData.Range("A1").CurrentRegion.AdvancedFilter _
        Action:=xlFilterCopy, _
        CopyToRange:=wsResult.Range("A1"), _
        Unique:=True

    wsResult.Columns.AutoFit
    wsResult.Activate
    MsgBox "已提取不重複記錄，結果在「唯一記錄結果」工作表。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 入口：提取單一欄位的唯一值清單
Public Sub ExtractUniqueColumnValuesExample()
    On Error GoTo ErrHandler

    Dim wsData   As Worksheet
    Dim wsResult As Worksheet

    Set wsData = GetOrCreateWsUnique(ThisWorkbook, "去重來源資料")
    Call FillDuplicateData(wsData)

    Set wsResult = GetOrCreateWsUnique(ThisWorkbook, "部門唯一清單")

    ' 只針對 B 欄（部門）提取唯一值
    wsData.Range("B1:B9").AdvancedFilter _
        Action:=xlFilterCopy, _
        CopyToRange:=wsResult.Range("A1"), _
        Unique:=True

    wsResult.Columns.AutoFit
    wsResult.Activate
    MsgBox "已提取部門唯一清單，結果在「部門唯一清單」工作表。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 填入含重複列的測試資料
Private Sub FillDuplicateData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("姓名", "部門", "年資")
    ws.Range("A2:C2").Value = Array("王小明", "業務部", 3)
    ws.Range("A3:C3").Value = Array("李美玲", "工程部", 5)
    ws.Range("A4:C4").Value = Array("王小明", "業務部", 3)   ' 重複
    ws.Range("A5:C5").Value = Array("張志豪", "業務部", 7)
    ws.Range("A6:C6").Value = Array("陳雅婷", "工程部", 2)
    ws.Range("A7:C7").Value = Array("李美玲", "工程部", 5)   ' 重複
    ws.Range("A8:C8").Value = Array("林建宏", "人資部", 4)
    ws.Range("A9:C9").Value = Array("林建宏", "人資部", 4)   ' 重複
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表並清空
Private Function GetOrCreateWsUnique(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateWsUnique = ws
End Function