Attribute VB_Name = "NumberRangeGroupPivot"
Option Explicit
'*************************************************************************************
'模組名稱: NumberRangeGroupPivot
'功能說明: 數值欄位自動區間分組示範
'          將數值欄位（如：單價）依指定區間自動分組
'          例如：0~999、1000~1999、2000~2999 等
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/6
'
'*************************************************************************************

' 程式進入點
Sub TestNumberRangeGroupPivot()
    Call CreateNumberRangeGroupPivot
End Sub

' 建立數值區間分組樞紐分析表
Sub CreateNumberRangeGroupPivot()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim dataRange As Range
    Dim pf As PivotField

    Set wsData = GetOrCreateSheet(ThisWorkbook, "訂單明細")
    Call FillOrderData(wsData)
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "數值區間群組樞紐")

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="數值區間群組樞紐分析表")

    ' 列欄位：單價（數值，後續群組）
    With pt.PivotFields("單價")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 值欄位：訂單數
    pt.AddDataField pt.PivotFields("訂單數"), "訂單數合計", xlSum

    ' 對「單價」欄位進行數值區間自動分組
    ' 起始值 0、結束值 5000、每段 1000
    Set pf = pt.PivotFields("單價")
    pf.AutoGroup

    ' 或使用手動指定區間
    ' pf.DataRange.Group Start:=0, End:=5000, By:=1000

    wsPivot.Columns("A:C").AutoFit
    wsPivot.Activate
    MsgBox "數值區間群組樞紐分析表已建立！" & Chr(13) & _
           "單價欄位已依數值範圍自動分組。", vbInformation, "完成"
End Sub

' 填入訂單明細資料
Private Sub FillOrderData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "訂單編號"
    ws.Range("B1").Value = "單價"
    ws.Range("C1").Value = "訂單數"

    ws.Range("A2").Value = "ORD001" : ws.Range("B2").Value = 250   : ws.Range("C2").Value = 5
    ws.Range("A3").Value = "ORD002" : ws.Range("B3").Value = 1200  : ws.Range("C3").Value = 3
    ws.Range("A4").Value = "ORD003" : ws.Range("B4").Value = 850   : ws.Range("C4").Value = 8
    ws.Range("A5").Value = "ORD004" : ws.Range("B5").Value = 3500  : ws.Range("C5").Value = 2
    ws.Range("A6").Value = "ORD005" : ws.Range("B6").Value = 480   : ws.Range("C6").Value = 12
    ws.Range("A7").Value = "ORD006" : ws.Range("B7").Value = 2200  : ws.Range("C7").Value = 4
    ws.Range("A8").Value = "ORD007" : ws.Range("B8").Value = 990   : ws.Range("C8").Value = 6
    ws.Range("A9").Value = "ORD008" : ws.Range("B9").Value = 4800  : ws.Range("C9").Value = 1
    ws.Range("A10").Value = "ORD009" : ws.Range("B10").Value = 1750 : ws.Range("C10").Value = 9
    ws.Range("A11").Value = "ORD010" : ws.Range("B11").Value = 320  : ws.Range("C11").Value = 15

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
