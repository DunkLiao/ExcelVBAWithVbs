Attribute VB_Name = "ProjectStatusFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ProjectStatusFormatting
'功能說明: 依專案狀態欄位套用條件式格式，自動以顏色區分進行中、已完成、逾期等狀態
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/11
'
'*************************************************************************************

' 範例使用入口
Sub TestProjectStatusFormatting()
    Call CreateProjectStatusFormatting("專案狀態追蹤")
End Sub

' 建立專案狀態追蹤工作表並套用條件式格式
' sheetName: 工作表名稱
Sub CreateProjectStatusFormatting(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim statusRange As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    Call FillProjectData(ws)

    Set dataRange = ws.Range("A2:E11")
    Set statusRange = ws.Range("E2:E11")

    Call ApplyProjectStatusCF(dataRange, statusRange, ws)

    ws.Columns("A:E").AutoFit
    MsgBox "專案狀態條件式格式設定完成！", vbInformation, "完成"
End Sub

' 套用條件式格式 — 使用 RGB 底色與字色區分三種狀態
Private Sub ApplyProjectStatusCF( _
    ByVal dataRange As Range, _
    ByVal statusRange As Range, _
    ByVal ws As Worksheet)

    Dim fc As FormatCondition
    Dim statusInProgress As String
    Dim statusDone As String
    Dim statusOverdue As String

    statusInProgress = "進行中"
    statusDone = "已完成"
    statusOverdue = "逾期"

    statusRange.FormatConditions.Delete

    ' 狀態：進行中 — 藍色底
    Set fc = statusRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="=""" & statusInProgress & """")
    fc.Interior.Color = RGB(173, 216, 230)
    fc.Font.Color = RGB(0, 70, 127)
    fc.Font.Bold = True

    ' 狀態：已完成 — 綠色底
    Set fc = statusRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="=""" & statusDone & """")
    fc.Interior.Color = RGB(144, 238, 144)
    fc.Font.Color = RGB(0, 100, 0)
    fc.Font.Bold = True

    ' 狀態：逾期 — 紅色底
    Set fc = statusRange.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:="=""" & statusOverdue & """")
    fc.Interior.Color = RGB(255, 182, 193)
    fc.Font.Color = RGB(139, 0, 0)
    fc.Font.Bold = True
End Sub

' 填入專案追蹤範例資料
Private Sub FillProjectData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "專案編號"
    ws.Range("B1").Value = "專案名稱"
    ws.Range("C1").Value = "負責人"
    ws.Range("D1").Value = "截止日期"
    ws.Range("E1").Value = "狀態"
    ws.Range("A1:E1").Font.Bold = True
    ws.Range("A1:E1").Interior.Color = RGB(70, 130, 180)
    ws.Range("A1:E1").Font.Color = RGB(255, 255, 255)

    ws.Range("A2").Value = "P001": ws.Range("B2").Value = "系統升級":     ws.Range("C2").Value = "王大明": ws.Range("D2").Value = "2026/03/31": ws.Range("E2").Value = "已完成"
    ws.Range("A3").Value = "P002": ws.Range("B3").Value = "報表自動化":   ws.Range("C3").Value = "林小玲": ws.Range("D3").Value = "2026/04/15": ws.Range("E3").Value = "進行中"
    ws.Range("A4").Value = "P003": ws.Range("B4").Value = "資料庫遷移":   ws.Range("C4").Value = "陳志明": ws.Range("D4").Value = "2026/02/28": ws.Range("E4").Value = "逾期"
    ws.Range("A5").Value = "P004": ws.Range("B5").Value = "前端改版":     ws.Range("C5").Value = "張雅惠": ws.Range("D5").Value = "2026/05/01": ws.Range("E5").Value = "進行中"
    ws.Range("A6").Value = "P005": ws.Range("B6").Value = "API 整合":     ws.Range("C6").Value = "王大明": ws.Range("D6").Value = "2026/04/30": ws.Range("E6").Value = "已完成"
    ws.Range("A7").Value = "P006": ws.Range("B7").Value = "資安稽核":     ws.Range("C7").Value = "林小玲": ws.Range("D7").Value = "2026/03/15": ws.Range("E7").Value = "逾期"
    ws.Range("A8").Value = "P007": ws.Range("B8").Value = "ERP 導入":     ws.Range("C8").Value = "陳志明": ws.Range("D8").Value = "2026/06/30": ws.Range("E8").Value = "進行中"
    ws.Range("A9").Value = "P008": ws.Range("B9").Value = "效能優化":     ws.Range("C9").Value = "張雅惠": ws.Range("D9").Value = "2026/05/15": ws.Range("E9").Value = "進行中"
    ws.Range("A10").Value = "P009": ws.Range("B10").Value = "文件整理":   ws.Range("C10").Value = "王大明": ws.Range("D10").Value = "2026/04/01": ws.Range("E10").Value = "已完成"
    ws.Range("A11").Value = "P010": ws.Range("B11").Value = "使用者訓練": ws.Range("C11").Value = "林小玲": ws.Range("D11").Value = "2026/03/20": ws.Range("E11").Value = "逾期"
End Sub