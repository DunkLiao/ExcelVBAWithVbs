Attribute VB_Name = "PivotLineChartExample"
Option Explicit

' ============================================================
' 範例：以 VBA 建立樞紐折線圖（PivotChart - Line）
' 功能：基於現有樞紐分析表建立折線樞紐圖，並設定標題
' ============================================================
Sub CreatePivotLineChart()
    Dim ws      As Worksheet
    Dim pt      As PivotTable
    Dim chObj   As ChartObject
    Dim ch      As Chart

    On Error GoTo ErrHandler

    ' --- 確認作用中工作表有樞紐分析表 ---
    Set ws = ActiveSheet
    If ws.PivotTables.Count = 0 Then
        MsgBox "作用中工作表沒有樞紐分析表，請先切換至含樞紐分析表的工作表。", vbExclamation
        Exit Sub
    End If

    Set pt = ws.PivotTables(1)

    ' --- 插入樞紐折線圖 ---
    Set chObj = ws.ChartObjects.Add(Left:=20, Top:=200, Width:=450, Height:=280)
    Set ch = chObj.Chart

    ch.SetSourceData Source:=pt.TableRange1
    ch.ChartType = xlLine

    ' --- 設定標題 ---
    ch.HasTitle = True
    ch.ChartTitle.Text = pt.Name & " 趨勢折線圖"

    ' --- 設定圖例 ---
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionBottom

    MsgBox "樞紐折線圖已建立於工作表：" & ws.Name, vbInformation
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
