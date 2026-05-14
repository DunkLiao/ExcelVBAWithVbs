Attribute VB_Name = "PivotFromExternalData"
Option Explicit
'*************************************************************************************
'模組名稱: PivotFromExternalData
'功能說明: 從外部活頁簿讀取資料，在目前活頁簿建立樞紐分析表的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點：選擇外部 Excel 檔案後建立樞紐
Sub TestPivotFromExternalData()
    Dim srcPath As String
    srcPath = GetExternalFilePath()
    If srcPath = "" Then
        MsgBox "未選擇檔案，作業取消。", vbInformation, "取消"
        Exit Sub
    End If
    Call BuildPivotFromExternal(srcPath, "外部資料樞紐")
End Sub

' 從外部 Excel 讀取第一個工作表資料並建立樞紐
' srcPath: 外部活頁簿路徑
' pivotSheetName: 樞紐輸出工作表名稱
Sub BuildPivotFromExternal(ByVal srcPath As String, ByVal pivotSheetName As String)
    On Error GoTo ErrorHandler

    Dim srcWb      As Workbook
    Dim srcWs      As Worksheet
    Dim destWs     As Worksheet
    Dim tempWs     As Worksheet
    Dim pc         As PivotCache
    Dim pt         As PivotTable
    Dim srcLastRow As Long
    Dim srcLastCol As Long
    Dim srcRange   As Range

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set srcWb = Workbooks.Open(srcPath, ReadOnly:=True)
    Set srcWs = srcWb.Worksheets(1)

    srcLastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row
    srcLastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column

    If srcLastRow < 2 Or srcLastCol < 1 Then
        srcWb.Close SaveChanges:=False
        MsgBox "外部檔案沒有足夠的資料列。", vbExclamation, "提示"
        Exit Sub
    End If

    Set tempWs = GetOrCreateExternalSheet("_ExternalTemp")
    tempWs.Cells.Clear
    srcWs.Range(srcWs.Cells(1, 1), srcWs.Cells(srcLastRow, srcLastCol)).Copy _
        Destination:=tempWs.Cells(1, 1)

    srcWb.Close SaveChanges:=False

    Set srcRange = tempWs.Range(tempWs.Cells(1, 1), tempWs.Cells(srcLastRow, srcLastCol))
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=srcRange)

    Set destWs = GetOrCreateExternalSheet(pivotSheetName)
    destWs.Cells.Clear

    Set pt = pc.CreatePivotTable( _
        TableDestination:=destWs.Range("A3"), _
        TableName:="ExternalPivot")

    With pt
        .PivotFields(1).Orientation = xlRowField
        .PivotFields(1).Position = 1
        If srcLastCol >= 2 Then
            With .PivotFields(srcLastCol)
                .Orientation = xlDataField
                .Function = xlSum
                .NumberFormat = "#,##0"
            End With
        End If
        .TableStyle2 = "PivotStyleMedium9"
        .RowAxisLayout xlOutlineRow
    End With

    destWs.Columns.AutoFit

    Application.DisplayAlerts = False
    tempWs.Delete
    Application.DisplayAlerts = True

    Application.ScreenUpdating = True
    MsgBox "外部資料樞紐分析表已建立完成！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not srcWb Is Nothing Then
        On Error Resume Next
        srcWb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    MsgBox "建立外部資料樞紐時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 取得外部 Excel 檔案路徑
Private Function GetExternalFilePath() As String
    Dim dialog As Object
    Set dialog = Application.FileDialog(msoFileDialogFilePicker)
    dialog.Title = "請選擇外部 Excel 資料來源"
    dialog.Filters.Clear
    dialog.Filters.Add "Excel 活頁簿", "*.xlsx;*.xls;*.xlsm"
    If dialog.Show = -1 Then
        GetExternalFilePath = dialog.SelectedItems(1)
    Else
        GetExternalFilePath = ""
    End If
End Function

' 取得或建立工作表
Private Function GetOrCreateExternalSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateExternalSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateExternalSheet Is Nothing Then
        Set GetOrCreateExternalSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateExternalSheet.Name = sheetName
    End If
End Function
