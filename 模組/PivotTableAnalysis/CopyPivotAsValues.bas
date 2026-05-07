Attribute VB_Name = "CopyPivotAsValues"
Option Explicit
'*************************************************************************************
'模組名稱: CopyPivotAsValues
'功能說明: 將樞紐分析表結果複製為純值到新工作表，
'          避免樞紐分析表連結，方便獨立使用或列印
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/4
'
'*************************************************************************************

' 測試入口
Sub TestCopyPivotAsValues()
    Call CreateAndCopyPivotAsValues
End Sub

' 建立樞紐分析表並將結果複製為純值到新工作表
Sub CreateAndCopyPivotAsValues()
    Dim wsData      As Worksheet
    Dim wsPivot     As Worksheet
    Dim wsCopy      As Worksheet
    Dim pc          As PivotCache
    Dim pt          As PivotTable
    Dim dataRange   As Range
    Dim ptFullRange As Range

    ' 準備資料工作表
    Set wsData = GetOrCreateSheet(ThisWorkbook, "銷售資料")
    Call FillSalesData(wsData)

    ' 準備樞紐工作表
    Set wsPivot = GetOrCreateSheet(ThisWorkbook, "原始樞紐")

    Set dataRange = wsData.Range("A1").CurrentRegion

    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange)
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="複製來源樞紐")

    ' 列欄位: 地區
    With pt.PivotFields("地區")
        .Orientation = xlRowField
        .Position = 1
    End With

    ' 欄欄位: 產品
    With pt.PivotFields("產品")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ' 值欄位: 銷售額加總
    pt.AddDataField pt.PivotFields("銷售額"), "銷售額加總", xlSum
    ' 值欄位: 數量加總
    pt.AddDataField pt.PivotFields("數量"), "數量加總", xlSum

    ' 複製樞紐完整範圍
    Set ptFullRange = pt.TableRange2
    ptFullRange.Copy

    ' 準備目標工作表（純值）
    Set wsCopy = GetOrCreateSheet(ThisWorkbook, "純值結果")
    wsCopy.Range("A1").PasteSpecial Paste:=xlPasteValues
    wsCopy.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    ' 自動調整欄寬
    wsCopy.Columns.AutoFit

    wsCopy.Activate
    MsgBox "樞紐分析表已複製為純值到「純值結果」工作表！" & Chr(13) & _
           "此工作表無樞紐連結，可獨立編輯或列印。", _
           vbInformation, "完成"
End Sub

' 填入範例銷售資料（8筆：日期、地區、產品、數量、單價、銷售額）
Private Sub FillSalesData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "日期"
    ws.Range("B1").Value = "地區"
    ws.Range("C1").Value = "產品"
    ws.Range("D1").Value = "數量"
    ws.Range("E1").Value = "單價"
    ws.Range("F1").Value = "銷售額"

    ws.Range("A2").Value = DateSerial(2026, 1, 5)
    ws.Range("B2").Value = "北區"
    ws.Range("C2").Value = "產品A"
    ws.Range("D2").Value = 100
    ws.Range("E2").Value = 50
    ws.Range("F2").Value = 5000

    ws.Range("A3").Value = DateSerial(2026, 1, 10)
    ws.Range("B3").Value = "南區"
    ws.Range("C3").Value = "產品B"
    ws.Range("D3").Value = 80
    ws.Range("E3").Value = 75
    ws.Range("F3").Value = 6000

    ws.Range("A4").Value = DateSerial(2026, 2, 5)
    ws.Range("B4").Value = "北區"
    ws.Range("C4").Value = "產品B"
    ws.Range("D4").Value = 60
    ws.Range("E4").Value = 75
    ws.Range("F4").Value = 4500

    ws.Range("A5").Value = DateSerial(2026, 2, 15)
    ws.Range("B5").Value = "中區"
    ws.Range("C5").Value = "產品A"
    ws.Range("D5").Value = 120
    ws.Range("E5").Value = 50
    ws.Range("F5").Value = 6000

    ws.Range("A6").Value = DateSerial(2026, 3, 1)
    ws.Range("B6").Value = "南區"
    ws.Range("C6").Value = "產品A"
    ws.Range("D6").Value = 90
    ws.Range("E6").Value = 50
    ws.Range("F6").Value = 4500

    ws.Range("A7").Value = DateSerial(2026, 3, 20)
    ws.Range("B7").Value = "中區"
    ws.Range("C7").Value = "產品B"
    ws.Range("D7").Value = 150
    ws.Range("E7").Value = 75
    ws.Range("F7").Value = 11250

    ws.Range("A8").Value = DateSerial(2026, 4, 5)
    ws.Range("B8").Value = "北區"
    ws.Range("C8").Value = "產品A"
    ws.Range("D8").Value = 200
    ws.Range("E8").Value = 50
    ws.Range("F8").Value = 10000

    ws.Range("A9").Value = DateSerial(2026, 4, 18)
    ws.Range("B9").Value = "南區"
    ws.Range("C9").Value = "產品B"
    ws.Range("D9").Value = 110
    ws.Range("E9").Value = 75
    ws.Range("F9").Value = 8250

    ws.Columns("A").NumberFormat = "yyyy/m/d"
    ws.Columns("A:F").AutoFit
End Sub

' 取得或建立工作表，並清除現有內容
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
