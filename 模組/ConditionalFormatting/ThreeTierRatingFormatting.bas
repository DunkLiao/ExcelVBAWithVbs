Attribute VB_Name = "ThreeTierRatingFormatting"
Option Explicit

'*************************************************************************************
'模組名稱: ThreeTierRatingFormatting
'功能說明: 依 A/B/C 評等設定三段條件格式化背景色
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub ApplyThreeTierRatingFormat()
    '對選取範圍依 A/B/C 評等套用不同背景色條件格式化
    Dim rng As Range
    Dim fc As FormatCondition

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取含有 A/B/C 評等的儲存格範圍！", vbExclamation
        Exit Sub
    End If

    Set rng = Selection
    rng.FormatConditions.Delete

    ' A 評等 - 綠色
    Set fc = rng.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:=""A"")
    fc.Interior.Color = RGB(146, 208, 80)
    fc.Font.Bold = True

    ' B 評等 - 黃色
    Set fc = rng.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:=""B"")
    fc.Interior.Color = RGB(255, 217, 102)
    fc.Font.Bold = False

    ' C 評等 - 紅色
    Set fc = rng.FormatConditions.Add( _
        Type:=xlCellValue, _
        Operator:=xlEqual, _
        Formula1:=""C"")
    fc.Interior.Color = RGB(255, 102, 102)
    fc.Font.Bold = True

    MsgBox "A/B/C 評等條件格式化已套用！", vbInformation
End Sub
