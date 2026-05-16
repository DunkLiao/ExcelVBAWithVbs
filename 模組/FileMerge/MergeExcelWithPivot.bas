Attribute VB_Name = "MergeExcelWithPivot"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelWithPivot
'功能說明: 合併指定資料夾內所有 .xlsx 的第一張工作表資料至主表，
'          合併完成後自動在新工作表建立樞紐分析表進行彙總
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

Sub MergeExcelFilesAndCreatePivot()
    Dim folderPath  As String
    Dim fileName    As String
    Dim wbSrc       As Workbook
    Dim wsDest      As Worksheet
    Dim wsPivot     As Worksheet
    Dim wsSrc       As Worksheet
    Dim destRow     As Long
    Dim srcLastRow  As Long
    Dim srcLastCol  As Long
    Dim hasHeader   As Boolean

    ' 選擇資料夾
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇要合併的 Excel 檔案資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation
            Exit Sub
        End If
        folderPath = .SelectedItems(1)
    End With

    If Right(folderPath, 1) <> "" Then folderPath = folderPath & ""

    ' 建立或清空合併目標工作表
    Set wsDest = GetOrCreateMergeSheet(ThisWorkbook, "合併資料")
    wsDest.Cells.Clear
    destRow = 1
    hasHeader = False

    ' 逐一合併
    fileName = Dir(folderPath & "*.xlsx")
    Do While fileName <> ""
        If fileName <> ThisWorkbook.Name Then
            Application.ScreenUpdating = False
            Set wbSrc = Workbooks.Open(folderPath & fileName, ReadOnly:=True)
            Set wsSrc = wbSrc.Worksheets(1)

            srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
            srcLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

            If srcLastRow >= 1 Then
                If Not hasHeader Then
                    ' 第一個檔案：複製含標題
                    wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(srcLastRow, srcLastCol)) _
                         .Copy wsDest.Cells(destRow, 1)
                    destRow = destRow + srcLastRow
                    hasHeader = True
                Else
                    ' 後續檔案：跳過標題列（第1行）
                    If srcLastRow > 1 Then
                        wsSrc.Range(wsSrc.Cells(2, 1), wsSrc.Cells(srcLastRow, srcLastCol)) _
                             .Copy wsDest.Cells(destRow, 1)
                        destRow = destRow + srcLastRow - 1
                    End If
                End If
            End If

            wbSrc.Close SaveChanges:=False
        End If
        fileName = Dir
    Loop

    Application.ScreenUpdating = True

    If destRow = 1 Then
        MsgBox "未找到任何 .xlsx 檔案。", vbExclamation
        Exit Sub
    End If

    ' 自動調整欄寬
    wsDest.Columns.AutoFit

    ' 建立樞紐分析表
    Call BuildSummaryPivot(wsDest, ThisWorkbook)

    MsgBox "合併完成，共 " & (destRow - 2) & " 筆資料，樞紐分析表已建立！", vbInformation, "完成"
End Sub

' 在新工作表建立基本樞紐分析表（以第1欄為列，第2欄為值）
Private Sub BuildSummaryPivot(ByVal wsSrc As Worksheet, ByVal wb As Workbook)
    Dim pc          As PivotCache
    Dim pt          As PivotTable
    Dim wsPivot     As Worksheet
    Dim srcLastRow  As Long
    Dim srcLastCol  As Long
    Dim rngSrc      As Range
    Dim col1Name    As String
    Dim col2Name    As String

    srcLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    srcLastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column

    If srcLastRow < 2 Or srcLastCol < 2 Then Exit Sub

    Set rngSrc = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(srcLastRow, srcLastCol))
    col1Name = CStr(wsSrc.Cells(1, 1).Value)
    col2Name = CStr(wsSrc.Cells(1, 2).Value)

    ' 移除既有樞紐工作表
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("合併樞紐")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    Set wsPivot = wb.Worksheets.Add
    wsPivot.Name = "合併樞紐"

    Set pc = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=rngSrc)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="合併樞紐表")

    With pt
        .PivotFields(col1Name).Orientation = xlRowField
        .PivotFields(col1Name).Position = 1
        If IsNumeric(wsSrc.Cells(2, 2).Value) Then
            .PivotFields(col2Name).Orientation = xlDataField
            .PivotFields(col2Name).Function = xlSum
            .PivotFields(col2Name).NumberFormat = "#,##0"
        End If
        .TableStyle2 = "PivotStyleMedium9"
    End With

    wsPivot.Columns.AutoFit
End Sub

Private Function GetOrCreateMergeSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateMergeSheet = ws
End Function
