Attribute VB_Name = "SetSheetPureText"
Option Explicit
'*************************************************************************************
'專案名稱: 風管系統
'功能描述: 設定excel報表只貼上值(for監控報表)
'https://zh-tw.extendoffice.com/documents/excel/4140-excel-save-workbook-as-values.html
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期：2024/7/10
'
'改版日期:
'改版備註: 2024/7/11 增加設定所有Print Rage
'
'*************************************************************************************

'只設定值(PringRange)
Sub S01_設定excel報表只貼上值PringRange()
    With ActiveSheet
            If .PageSetup.PrintArea <> "" Then
                .Range(.PageSetup.PrintArea).Copy
                .Range(.PageSetup.PrintArea).PasteSpecial xlPasteValues
            Else
                msgbox "設定excel報表只貼上值(PringRange)無設定列印範圍!"
                Exit Sub
            End If
    End With
     Application.CutCopyMode = False
     msgbox "設定excel報表只貼上值(PringRange)OK!"
End Sub
'只設定值
Sub S02_設定excel報表只貼上值()
    With ActiveSheet
            .Cells.Copy
            .Cells.PasteSpecial xlPasteValues
    End With
    Application.CutCopyMode = False
    msgbox "設定excel報表只貼上值OK!"
End Sub

'設定活頁簿所有工作頁內容只有值
Sub S03_設定活頁簿所有工作頁內容只有值()
    Dim wsh As Worksheet
    For Each wsh In ActiveWorkbook.Worksheets
        With wsh
            .Cells.Copy
            .Cells.PasteSpecial xlPasteValues
        End With
    Next
    Set wsh = Nothing
    Application.CutCopyMode = False
    msgbox "設定活頁簿所有工作頁內容只有值OK!"
End Sub

'設定活頁簿所有工作頁內容只有值PringRange
Sub S03_設定活頁簿所有工作頁內容只有值PringRange()
    Dim wsh As Worksheet
    For Each wsh In ActiveWorkbook.Worksheets
        With wsh
            If .PageSetup.PrintArea <> "" Then
                .Range(.PageSetup.PrintArea).Copy
                .Range(.PageSetup.PrintArea).PasteSpecial xlPasteValues
            End If
        End With
    Next
    Set wsh = Nothing
    Application.CutCopyMode = False
    msgbox "設定活頁簿所有工作頁內容只有值PringRangeOK!"
End Sub
