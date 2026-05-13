Attribute VB_Name = "ProgressBarFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: 進度條格式
'功能說明: 使用條件式格式將儲存格顯示為百分比進度條樣式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub TestProgressBarFormatting()
    Call ApplyProgressBarFormatting("進度條格式範例")
End Sub

Sub ApplyProgressBarFormatting(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim targetRange As Range
    Dim cf As Databar

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear

    ws.Range("A1").Value = "專案名稱"
    ws.Range("B1").Value = "完成進度(%)"

    ws.Range("A2").Value = "UI設計"      : ws.Range("B2").Value = 85
    ws.Range("A3").Value = "後端開發"    : ws.Range("B3").Value = 60
    ws.Range("A4").Value = "資料庫建置" : ws.Range("B4").Value = 100
    ws.Range("A5").Value = "測試驗收"    : ws.Range("B5").Value = 30
    ws.Range("A6").Value = "文件撰寫"    : ws.Range("B6").Value = 45
    ws.Range("A7").Value = "部署上線"    : ws.Range("B7").Value = 10

    Set targetRange = ws.Range("B2:B7")
    targetRange.FormatConditions.Delete

    Set cf = targetRange.FormatConditions.AddDatabar
    cf.MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
    cf.MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=100
    cf.BarColor.Color = RGB(70, 180, 80)
    cf.BarFillType = xlDataBarFillGradient
    cf.ShowValue = True

    targetRange.NumberFormat = "0\%"
    ws.Columns("A:B").AutoFit
    ws.Range("A1:B1").Font.Bold = True

    MsgBox "進度條格式已套用完成！", vbInformation, "完成"
End Sub
