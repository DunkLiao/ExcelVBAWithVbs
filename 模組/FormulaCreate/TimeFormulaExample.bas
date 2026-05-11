Attribute VB_Name = "TimeFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: TimeFormulaExample
'功能說明: 在 Excel 中批次建立時間相關公式的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestTimeFormulas()
    Call CreateTimeFormulas("時間公式範例")
End Sub

' 建立時間相關公式範例
' sheetName: 要寫入的工作表名稱
Sub CreateTimeFormulas(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet

    Set ws = GetOrCreateWorksheetTime(sheetName)
    ws.Cells.Clear

    ' 標題列
    ws.Range("A1").Value = "說明"
    ws.Range("B1").Value = "公式結果"

    ' 目前時間
    ws.Range("A2").Value = "目前時間"
    ws.Range("B2").Formula = "=NOW()"
    ws.Range("B2").NumberFormat = "hh:mm:ss"

    ' 目前日期
    ws.Range("A3").Value = "目前日期"
    ws.Range("B3").Formula = "=TODAY()"
    ws.Range("B3").NumberFormat = "yyyy/mm/dd"

    ' 小時數
    ws.Range("A4").Value = "目前小時"
    ws.Range("B4").Formula = "=HOUR(NOW())"

    ' 分鐘數
    ws.Range("A5").Value = "目前分鐘"
    ws.Range("B5").Formula = "=MINUTE(NOW())"

    ' 秒數
    ws.Range("A6").Value = "目前秒數"
    ws.Range("B6").Formula = "=SECOND(NOW())"

    ' 工時計算：上班打卡到下班打卡
    ws.Range("A8").Value = "上班時間"
    ws.Range("B8").Value = "09:00"
    ws.Range("B8").NumberFormat = "hh:mm"

    ws.Range("A9").Value = "下班時間"
    ws.Range("B9").Value = "18:30"
    ws.Range("B9").NumberFormat = "hh:mm"

    ws.Range("A10").Value = "工作時數（小時）"
    ws.Range("B10").Formula = "=(B9-B8)*24"
    ws.Range("B10").NumberFormat = "0.00"

    ' 兩日期間隔天數
    ws.Range("A12").Value = "起始日期"
    ws.Range("B12").Value = "2026/1/1"
    ws.Range("B12").NumberFormat = "yyyy/mm/dd"

    ws.Range("A13").Value = "結束日期"
    ws.Range("B13").Formula = "=TODAY()"
    ws.Range("B13").NumberFormat = "yyyy/mm/dd"

    ws.Range("A14").Value = "間隔天數"
    ws.Range("B14").Formula = "=B13-B12"
    ws.Range("B14").NumberFormat = "0"

    ' 月份差
    ws.Range("A15").Value = "間隔月數 (DATEDIF)"
    ws.Range("B15").Formula = "=DATEDIF(B12,B13,""M"")"

    ' 當月最後一天
    ws.Range("A17").Value = "當月最後一天"
    ws.Range("B17").Formula = "=EOMONTH(TODAY(),0)"
    ws.Range("B17").NumberFormat = "yyyy/mm/dd"

    ' 下個月第一天
    ws.Range("A18").Value = "下個月第一天"
    ws.Range("B18").Formula = "=EOMONTH(TODAY(),0)+1"
    ws.Range("B18").NumberFormat = "yyyy/mm/dd"

    ' 當週星期幾
    ws.Range("A19").Value = "今天是星期幾（1=週日）"
    ws.Range("B19").Formula = "=WEEKDAY(TODAY(),1)"

    ws.Range("A1").Font.Bold = True
    ws.Range("B1").Font.Bold = True
    ws.Columns("A:B").AutoFit

    MsgBox "時間公式範例已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立時間公式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateWorksheetTime(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateWorksheetTime = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateWorksheetTime Is Nothing Then
        Set GetOrCreateWorksheetTime = ThisWorkbook.Worksheets.Add
        GetOrCreateWorksheetTime.Name = sheetName
    End If
End Function
