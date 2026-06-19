Attribute VB_Name = "ClearAllExceptData"
Option Explicit
'*************************************************************************************
'模組名稱: ClearAllExceptData
'功能說明: 清除工作表中所有格式、註解、驗證等內容，僅保留儲存格資料值
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestClearAllExceptData()
    Call ClearAllExceptDataValues
End Sub

Sub ClearAllExceptDataValues()
    Dim ws As Worksheet
    Dim usedRange As Range
    Dim cell As Range
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim msg As String
    msg = "此操作將清除目前工作表中所有格式、註解、驗證與條件式格式，" & vbCrLf & _
          "僅保留資料值。是否繼續？"
    
    If MsgBox(msg, vbYesNo + vbQuestion, "確認清除") <> vbYes Then
        MsgBox "已取消操作", vbInformation, "取消"
        Exit Sub
    End If
    
    Set ws = ActiveSheet
    
    ' 移除條件式格式
    ws.Cells.FormatConditions.Delete
    
    ' 清除資料驗證
    ws.Cells.Validation.Delete
    
    ' 移除所有註解
    ws.Cells.ClearComments
    
    ' 移除超連結
    ws.Hyperlinks.Delete
    
    ' 清除所有圖表
    Dim chtObj As ChartObject
    For Each chtObj In ws.ChartObjects
        chtObj.Delete
    Next chtObj
    
    ' 清除所有圖片與圖形
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    
    ' 清除所有樞紐分析表
    Dim pt As PivotTable
    For Each pt In ws.PivotTables
        pt.TableRange2.Clear
    Next pt
    
    ' 清除格式但保留值
    Set usedRange = ws.UsedRange
    usedRange.ClearFormats
    
    ' 重設欄寬與列高為預設值
    usedRange.EntireColumn.ColumnWidth = 8.43
    usedRange.EntireRow.RowHeight = 15
    
    ' 移除合併儲存格
    usedRange.UnMerge
    
    ' 移除大綱（群組）
    ws.Cells.ClearOutline
    
    ' 移除自動篩選
    If ws.AutoFilterMode Then
        ws.AutoFilterMode = False
    End If
    
    Application.ScreenUpdating = True
    MsgBox "清除完成！所有格式已移除，僅保留資料值。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "清除格式時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
