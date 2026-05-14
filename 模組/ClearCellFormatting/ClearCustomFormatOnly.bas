Attribute VB_Name = "ClearCustomFormatOnly"
Option Explicit
'*************************************************************************************
'模組名稱: ClearCustomFormatOnly
'功能說明: 僅清除選取範圍或整個工作表中的自訂數字格式，還原為通用格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestClearCustomFormatOnly()
    Call CreateCustomFormatDemo
    Call ClearCustomNumberFormatInSheet(ActiveSheet)
End Sub

' 建立含自訂格式的示範資料
Private Sub CreateCustomFormatDemo()
    Dim ws As Worksheet
    Set ws = GetOrCreateCustFmtWs(ThisWorkbook, "自訂格式範例")
    ws.Cells.Clear
    ws.Activate

    ws.Range("A1").Value = "項目"
    ws.Range("B1").Value = "數值"
    ws.Range("C1").Value = "格式說明"
    ws.Range("A1:C1").Font.Bold = True

    ws.Range("A2").Value = "日期格式"
    ws.Range("B2").Value = Now
    ws.Range("B2").NumberFormat = "yyyy/mm/dd hh:mm"
    ws.Range("C2").Value = "yyyy/mm/dd hh:mm（自訂）"

    ws.Range("A3").Value = "貨幣格式"
    ws.Range("B3").Value = 12345.678
    ws.Range("B3").NumberFormat = "$#,##0.00"
    ws.Range("C3").Value = "$#,##0.00（自訂）"

    ws.Range("A4").Value = "百分比格式"
    ws.Range("B4").Value = 0.856
    ws.Range("B4").NumberFormat = "0.0%"
    ws.Range("C4").Value = "0.0%（自訂）"

    ws.Range("A5").Value = "科學記號"
    ws.Range("B5").Value = 0.0000123
    ws.Range("B5").NumberFormat = "0.00E+00"
    ws.Range("C5").Value = "0.00E+00（自訂）"

    ws.Range("A6").Value = "大數格式"
    ws.Range("B6").Value = 9876543
    ws.Range("B6").NumberFormat = "#,##0.0,,\"M\""
    ws.Range("C6").Value = "#,##0.0,,M（自訂，以百萬顯示）"

    ws.Columns("A:C").AutoFit
End Sub

' 清除工作表中所有自訂數字格式（還原為通用格式）
Sub ClearCustomNumberFormatInSheet(ByVal ws As Worksheet)
    On Error GoTo ErrorHandler

    Dim cell As Range
    Dim clearedCount As Long
    clearedCount = 0

    Dim standardFormats(0 To 5) As String
    standardFormats(0) = "General"
    standardFormats(1) = "@"
    standardFormats(2) = "0"
    standardFormats(3) = "#,##0"
    standardFormats(4) = "0.00"
    standardFormats(5) = "#,##0.00"

    Application.ScreenUpdating = False

    For Each cell In ws.UsedRange
        Dim fmt As String
        fmt = cell.NumberFormat
        Dim isStandard As Boolean
        isStandard = False
        Dim j As Long
        For j = 0 To 5
            If fmt = standardFormats(j) Then
                isStandard = True
                Exit For
            End If
        Next j
        If Not isStandard Then
            cell.NumberFormat = "General"
            clearedCount = clearedCount + 1
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "已清除 " & clearedCount & " 個自訂數字格式，已還原為通用格式。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清除格式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 清除選取範圍中的自訂數字格式
Sub ClearCustomNumberFormatInSelection()
    On Error GoTo ErrorHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取一個儲存格範圍。", vbExclamation, "提示"
        Exit Sub
    End If

    Dim rng As Range
    Set rng = Selection
    Dim cell As Range
    Dim clearedCount As Long
    clearedCount = 0

    Application.ScreenUpdating = False
    For Each cell In rng
        If cell.NumberFormat <> "General" And cell.NumberFormat <> "@" Then
            cell.NumberFormat = "General"
            clearedCount = clearedCount + 1
        End If
    Next cell
    Application.ScreenUpdating = True

    MsgBox "已清除選取範圍中 " & clearedCount & " 個自訂數字格式。", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清除格式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateCustFmtWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateCustFmtWs = ws
End Function
