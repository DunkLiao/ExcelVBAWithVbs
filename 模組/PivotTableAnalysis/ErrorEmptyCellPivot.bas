Attribute VB_Name = "ErrorEmptyCellPivot"
Option Explicit
'*************************************************************************************
'模組名稱: ErrorEmptyCellPivot
'功能說明: 空白及錯誤儲存格顯示設定示範
'          設定樞紐分析表在遇到空白資料時顯示指定文字
'          以及當計算發生錯誤時的替代顯示文字
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/6
'
'*************************************************************************************

' 程式進入點
Sub TestErrorEmptyCellPivot()
    Call CreateErrorEmptyCellPivot
End Sub

' 建立含空白及錯誤顯示設定的樞紐分析表
Sub CreateErrorEmptyCellPivot()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range

    Set wsData = GetOrCreateSheet(ThisWorkbook, "含空值銷售")
    Call FillSparseData(wsData)
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "空白錯誤樞紐")

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="空白錯誤樞紐分析表")

    ' 列欄位：地區
    With pt.PivotFields("地區")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 欄欄位：產品
    With pt.PivotFields("產品")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ' 值欄位：銷售額
    pt.AddDataField pt.PivotFields("銷售額"), "銷售額合計", xlSum

    ' 設定空白儲存格顯示文字
    pt.NullString = "－"

    ' 設定錯誤值顯示文字
    pt.ErrorString = "錯誤"
    pt.DisplayErrorString = True

    ' 開啟空白顯示
    pt.DisplayNullString = True

    wsPivot.Columns("A:E").AutoFit
    wsPivot.Activate
    MsgBox "空白/錯誤顯示樞紐分析表已建立！" & Chr(13) & _
           "空白儲存格顯示：『－』" & Chr(13) & _
           "錯誤值顯示：『錯誤』", vbInformation, "完成"
End Sub

' 填入含稀疏（空值）的銷售資料
Private Sub FillSparseData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "地區"
    ws.Range("B1").Value = "產品"
    ws.Range("C1").Value = "銷售額"

    ' 刻意讓部分地區/產品組合缺少資料（產生空白格）
    ws.Range("A2").Value = "北部" : ws.Range("B2").Value = "產品A" : ws.Range("C2").Value = 50000
    ws.Range("A3").Value = "北部" : ws.Range("B3").Value = "產品B" : ws.Range("C3").Value = 32000
    ws.Range("A4").Value = "南部" : ws.Range("B4").Value = "產品A" : ws.Range("C4").Value = 41000
    ' 南部 產品B 刻意省略（樞紐會顯示空白）
    ws.Range("A5").Value = "東部" : ws.Range("B5").Value = "產品C" : ws.Range("C5").Value = 28000
    ws.Range("A6").Value = "北部" : ws.Range("B6").Value = "產品C" : ws.Range("C6").Value = 19000
    ' 東部 產品A、產品B 刻意省略

    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表，並清除內容
Private Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheet = ws
End Function
