Option Explicit
Attribute VB_Name = "PivotDefaultFieldSettings"
'*************************************************************************************

'模組名稱: PivotDefaultFieldSettings

'功能說明: 建立樞紐分析表並示範設定各欄位的預設彙總方式與數字格式

'

'版權所有: Dunk

'程式設計: Dunk

'撰寫日期: 2026/5/27

'

'*************************************************************************************




Sub CreatePivotDefaultFieldSettings()

    Dim ws As Worksheet

    Dim wsPivot As Worksheet

    Dim pt As PivotTable

    Dim pc As PivotCache

    Dim dataRange As Range

    Dim lastRow As Long

    Dim lastCol As Integer



    Set ws = ActiveSheet

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column



    If lastRow < 2 Or lastCol < 2 Then

        MsgBox "請確認工作表有足夠的資料（至少兩欄、兩列）。", vbExclamation, "提示"

        Exit Sub

    End If



    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))



    ' 建立樞紐分析快取

    Set pc = ThisWorkbook.PivotCaches.Create( _

        SourceType:=xlDatabase, _

        SourceData:=dataRange)



    ' 建立新工作表放置樞紐

    On Error Resume Next

    Application.DisplayAlerts = False

    ThisWorkbook.Worksheets("PivotFieldSettings").Delete

    Application.DisplayAlerts = True

    On Error GoTo 0



    Set wsPivot = ThisWorkbook.Worksheets.Add( _

        After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

    wsPivot.Name = "PivotFieldSettings"



    ' 建立樞紐分析表

    Set pt = pc.CreatePivotTable( _

        TableDestination:=wsPivot.Range("A3"), _

        TableName:="PivotFieldDemo")



    Application.ScreenUpdating = False



    With pt

        ' 設定列標籤（第一個文字欄）

        Dim fldRow As PivotField

        Set fldRow = .PivotFields(ws.Cells(1, 1).Value)

        fldRow.Orientation = xlRowField

        fldRow.Position = 1



        ' 設定欄標籤（第二個欄，若存在）

        If lastCol >= 3 Then

            Dim fldCol As PivotField

            Set fldCol = .PivotFields(ws.Cells(1, 2).Value)

            fldCol.Orientation = xlColumnField

            fldCol.Position = 1

        End If



        ' 設定數值欄（最後一欄）

        Dim fldVal As PivotField

        Set fldVal = .PivotFields(ws.Cells(1, lastCol).Value)

        fldVal.Orientation = xlDataField

        fldVal.Function = xlSum

        fldVal.NumberFormat = "#,##0.00"



        ' 設定樞紐樣式

        .TableStyle2 = "PivotStyleMedium9"



        ' 設定列配置

        .RowAxisLayout xlTabularRow

    End With



    wsPivot.Columns.AutoFit

    Application.ScreenUpdating = True



    MsgBox "樞紐分析表（含預設欄位設定）已建立完成！", vbInformation, "完成"

End Sub

