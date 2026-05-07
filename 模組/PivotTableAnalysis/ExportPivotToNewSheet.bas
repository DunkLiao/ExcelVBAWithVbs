Attribute VB_Name = "ExportPivotToNewSheet"
Option Explicit

' ============================================================
' 範例：將樞紐分析表的數值結果複製貼上至新工作表（去除公式）
' 功能：選取作用中樞紐分析表，以「只貼值」方式匯出至新工作表
' ============================================================
Sub ExportPivotToNewSheet()
    Dim ws          As Worksheet
    Dim wsNew       As Worksheet
    Dim pt          As PivotTable
    Dim rngPivot    As Range

    On Error GoTo ErrHandler

    ' --- 確認作用中工作表有樞紐分析表 ---
    Set ws = ActiveSheet
    If ws.PivotTables.Count = 0 Then
        MsgBox "作用中工作表沒有樞紐分析表，請先切換至含樞紐分析表的工作表。", vbExclamation
        Exit Sub
    End If

    Set pt = ws.PivotTables(1)
    Set rngPivot = pt.TableRange2

    ' --- 建立新工作表並貼上純值 ---
    Set wsNew = ThisWorkbook.Worksheets.Add( _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsNew.Name = "樞紐匯出_" & Format(Now(), "mmddHHmm")

    rngPivot.Copy
    wsNew.Range("A1").PasteSpecial Paste:=xlPasteValues
    wsNew.Range("A1").PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    wsNew.Columns.AutoFit
    MsgBox "樞紐分析表已匯出至工作表：" & wsNew.Name, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
