Attribute VB_Name = "ManualGroupPivot"
Option Explicit
'*************************************************************************************
'模組名稱: ManualGroupPivot
'功能說明: 建立樞紐分析表並示範手動對文字項目進行群組設定的操作說明
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub CreateManualGroupPivot()
    On Error GoTo ErrHandler
    Dim wb      As Workbook
    Dim wsData  As Worksheet
    Dim wsPivot As Worksheet
    Dim pc      As PivotCache
    Dim pt      As PivotTable
    Dim pf      As PivotField

    Set wb = ThisWorkbook
    Set wsData  = GetOrCreateSheetGrp(wb, "手動分組來源資料")
    Call FillGroupSourceData(wsData)
    Set wsPivot = GetOrCreateSheetGrp(wb, "手動分組樞紐")

    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.Range("A1").CurrentRegion)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="手動分組PT")

    Set pf = pt.PivotFields("區域")
    pf.Orientation = xlRowField
    pf.Position     = 1

    With pt.PivotFields("銷售額")
        .Orientation  = xlDataField
        .Function     = xlSum
        .NumberFormat = "#,##0"
        .Name          = "總銷售額"
    End With

    wsPivot.Range("A1").Value = "提示：請手動在樞紐分析表中選取欲群組的項目，再按右鍵選擇群組"
    wsPivot.Range("A2").Value = "（VBA 自動群組僅支援數值與日期欄位；文字欄位需手動操作）"
    pt.TableStyle2 = "PivotStyleMedium9"
    wsPivot.Columns.AutoFit
    wsPivot.Activate
    MsgBox "手動分組樞紐分析表已建立！請在樞紐中選取多個項目後，按右鍵群組完成文字群組。", vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

Private Sub FillGroupSourceData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("區域", "業務員", "銷售額")
    ws.Range("A2:C2").Value = Array("北區", "張小明", 85000)
    ws.Range("A3:C3").Value = Array("北區", "李美華", 92000)
    ws.Range("A4:C4").Value = Array("中區", "王大同", 76000)
    ws.Range("A5:C5").Value = Array("中區", "陳志偉", 68000)
    ws.Range("A6:C6").Value = Array("南區", "林雅芳", 95000)
    ws.Range("A7:C7").Value = Array("南區", "吳俊宏", 81000)
    ws.Range("A8:C8").Value = Array("東區", "黃淑娟", 73000)
    ws.Range("A9:C9").Value = Array("東區", "蔡建國", 88000)
    ws.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreateSheetGrp(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetGrp = ws
End Function

