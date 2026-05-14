Attribute VB_Name = "CustomIconSetFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: CustomIconSetFormatting
'功能說明: 使用自訂圖示集條件式格式標示資料達成率（5 種圖示）
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/14
'
'*************************************************************************************

' 測試用入口
Sub TestCustomIconSetFormatting()
    Call ApplyCustomIconSetFormatting
End Sub

' 套用自訂圖示集條件式格式
Sub ApplyCustomIconSetFormatting()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateIconWs(ThisWorkbook, "圖示集格式範例")
    ws.Cells.Clear

    ' 填入測試資料（達成率）
    ws.Range("A1").Value = "業務員"
    ws.Range("B1").Value = "目標"
    ws.Range("C1").Value = "實績"
    ws.Range("D1").Value = "達成率%"
    ws.Range("A1:D1").Font.Bold = True

    Dim salesData(1 To 8, 1 To 3) As Variant
    salesData(1, 1) = "王小明" : salesData(1, 2) = 100000 : salesData(1, 3) = 120000
    salesData(2, 1) = "李大華" : salesData(2, 2) = 100000 : salesData(2, 3) = 95000
    salesData(3, 1) = "陳美玲" : salesData(3, 2) = 100000 : salesData(3, 3) = 110000
    salesData(4, 1) = "林俊傑" : salesData(4, 2) = 100000 : salesData(4, 3) = 68000
    salesData(5, 1) = "張志遠" : salesData(5, 2) = 100000 : salesData(5, 3) = 85000
    salesData(6, 1) = "吳雅婷" : salesData(6, 2) = 100000 : salesData(6, 3) = 130000
    salesData(7, 1) = "黃建宏" : salesData(7, 2) = 100000 : salesData(7, 3) = 45000
    salesData(8, 1) = "劉佳欣" : salesData(8, 2) = 100000 : salesData(8, 3) = 100000

    Dim i As Long
    For i = 1 To 8
        ws.Cells(i + 1, 1).Value = salesData(i, 1)
        ws.Cells(i + 1, 2).Value = salesData(i, 2)
        ws.Cells(i + 1, 3).Value = salesData(i, 3)
        ws.Cells(i + 1, 4).Formula = "=ROUND(C" & (i + 1) & "/B" & (i + 1) & "*100,1)"
    Next i

    ws.Columns("A:D").AutoFit

    ' 套用 5 格圖示集（達成率欄位 D2:D9）
    Dim rng As Range
    Set rng = ws.Range("D2:D9")
    rng.FormatConditions.Delete

    Dim ic As IconSetCondition
    Set ic = rng.FormatConditions.AddIconSetCondition()

    ic.IconSet = ws.Parent.IconSets(xl5Arrows)

    ' 自訂各段門檻（百分比型態）
    With ic.IconCriteria(1)
        .Type = xlConditionValueNumber
        .Value = 0
        .Operator = xlGreaterEqual
    End With

    With ic.IconCriteria(2)
        .Type = xlConditionValueNumber
        .Value = 60
        .Operator = xlGreaterEqual
    End With

    With ic.IconCriteria(3)
        .Type = xlConditionValueNumber
        .Value = 80
        .Operator = xlGreaterEqual
    End With

    With ic.IconCriteria(4)
        .Type = xlConditionValueNumber
        .Value = 100
        .Operator = xlGreaterEqual
    End With

    With ic.IconCriteria(5)
        .Type = xlConditionValueNumber
        .Value = 120
        .Operator = xlGreaterEqual
    End With

    ws.Activate
    MsgBox "自訂圖示集條件式格式已套用完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "套用圖示集格式時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得或建立工作表
Private Function GetOrCreateIconWs(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateIconWs = ws
End Function
