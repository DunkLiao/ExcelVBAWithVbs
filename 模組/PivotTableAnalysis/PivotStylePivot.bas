Attribute VB_Name = "PivotStylePivot"
Option Explicit
'*************************************************************************************
'模組名稱: PivotStylePivot
'功能說明: 樞紐分析表樣式套用示範
'          示範套用 Excel 內建 PivotTable 樣式
'          包含 Light / Medium / Dark 系列
'          並開啟帶狀列與標題列格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/6
'
'*************************************************************************************

' 程式進入點
Sub TestPivotStylePivot()
    Call CreatePivotWithStyles
End Sub

' 建立樞紐分析表並套用內建樣式
Sub CreatePivotWithStyles()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim df As PivotField

    Set wsData = GetOrCreateSheet(ThisWorkbook, "部門費用")
    Call FillDeptData(wsData)
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "樣式樞紐")

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="樣式樞紐分析表")

    ' 列欄位：部門
    With pt.PivotFields("部門")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 欄欄位：費用類型
    With pt.PivotFields("費用類型")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ' 值欄位：金額
    Set df = pt.AddDataField(pt.PivotFields("金額"), "金額合計", xlSum)
    df.NumberFormat = "#,##0"

    ' 套用內建樣式（Medium 藍色系）
    pt.TableStyle2 = "PivotStyleMedium2"

    ' 顯示帶狀列
    pt.ShowTableStyleColumnHeaders = True
    pt.ShowTableStyleRowHeaders = True
    pt.ShowTableStyleRowStripes = True

    ' 其他可用樣式（取消下方註解即可切換）
    ' pt.TableStyle2 = "PivotStyleLight16"   ' 淺色
    ' pt.TableStyle2 = "PivotStyleDark1"     ' 深色
    ' pt.TableStyle2 = "PivotStyleMedium9"   ' 橘色

    wsPivot.Columns("A:F").AutoFit
    wsPivot.Activate
    MsgBox "樣式樞紐分析表已建立！" & Chr(13) & _
           "目前套用：PivotStyleMedium2" & Chr(13) & _
           "可修改 TableStyle2 屬性切換其他樣式。", vbInformation, "完成"
End Sub

' 填入部門費用資料
Private Sub FillDeptData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "部門"
    ws.Range("B1").Value = "費用類型"
    ws.Range("C1").Value = "金額"

    ws.Range("A2").Value = "業務部" : ws.Range("B2").Value = "差旅費" : ws.Range("C2").Value = 25000
    ws.Range("A3").Value = "業務部" : ws.Range("B3").Value = "交際費" : ws.Range("C3").Value = 18000
    ws.Range("A4").Value = "業務部" : ws.Range("B4").Value = "辦公費" : ws.Range("C4").Value = 5000
    ws.Range("A5").Value = "研發部" : ws.Range("B5").Value = "差旅費" : ws.Range("C5").Value = 12000
    ws.Range("A6").Value = "研發部" : ws.Range("B6").Value = "辦公費" : ws.Range("C6").Value = 15000
    ws.Range("A7").Value = "研發部" : ws.Range("B7").Value = "設備費" : ws.Range("C7").Value = 45000
    ws.Range("A8").Value = "管理部" : ws.Range("B8").Value = "差旅費" : ws.Range("C8").Value = 8000
    ws.Range("A9").Value = "管理部" : ws.Range("B9").Value = "交際費" : ws.Range("C9").Value = 6000
    ws.Range("A10").Value = "管理部" : ws.Range("B10").Value = "辦公費" : ws.Range("C10").Value = 9000

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
