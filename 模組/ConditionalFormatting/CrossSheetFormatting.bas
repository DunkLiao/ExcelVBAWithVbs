Attribute VB_Name = "CrossSheetFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: CrossSheetFormatting
'功能說明: 根據另一個工作表的數值來設定條件式格式的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestCrossSheetFormatting()
    Call ApplyCrossSheetConditionalFormatting
End Sub

Sub ApplyCrossSheetConditionalFormatting()
    Dim wsTarget As Worksheet
    Dim wsRef As Worksheet
    Dim rng As Range
    Dim fc As FormatCondition

    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsTarget = ThisWorkbook.Worksheets("銷售統計")
    If Not wsTarget Is Nothing Then wsTarget.Delete
    Set wsRef = ThisWorkbook.Worksheets("目標基準")
    If Not wsRef Is Nothing Then wsRef.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 建立銷售統計工作表
    Set wsTarget = ThisWorkbook.Worksheets.Add
    wsTarget.Name = "銷售統計"

    wsTarget.Range("A1").Value = "業務員"
    wsTarget.Range("B1").Value = "實際銷售"
    wsTarget.Range("A1:B1").Font.Bold = True

    wsTarget.Range("A2").Value = "陳大文"
    wsTarget.Range("B2").Value = 92000

    wsTarget.Range("A3").Value = "林小美"
    wsTarget.Range("B3").Value = 68000

    wsTarget.Range("A4").Value = "黃建國"
    wsTarget.Range("B4").Value = 105000

    wsTarget.Range("A5").Value = "周美玲"
    wsTarget.Range("B5").Value = 77000

    ' 建立目標基準工作表
    Set wsRef = ThisWorkbook.Worksheets.Add
    wsRef.Name = "目標基準"
    wsRef.Range("A1").Value = "每月目標"
    wsRef.Range("B1").Value = 80000

    ' 定義範圍名稱，用於跨工作表參照
    ThisWorkbook.Names.Add _
        Name:="月目標", _
        RefersTo:="=目標基準!$B$1"

    ' 設定條件式格式：實際銷售 >= 月目標時標示綠色
    Set rng = wsTarget.Range("B2:B5")

    rng.FormatConditions.Delete

    Set fc = rng.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlGreaterEqual, _
        Formula1:="=月目標")

    fc.Interior.Color = RGB(198, 239, 206)
    fc.Font.Color = RGB(0, 97, 0)
    fc.Font.Bold = True

    MsgBox "跨工作表條件式格式已設定完成！" & vbCrLf & vbCrLf & _
           "實際銷售 >= 月目標（" & wsRef.Range("B1").Value & "）的儲存格將以綠色標示。", _
           vbInformation, "完成"
End Sub
