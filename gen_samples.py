# -*- coding: utf-8 -*-
"""
Generate one new sample .bas file per folder, CP950 encoding.
"""
import os

BASE = r"D:\VIbeCoding\ExcelVBAWithVbs\模組"

###############################################################################
# Helper: write a .bas file
###############################################################################
def write_bas(folder, filename, content):
    path = os.path.join(BASE, folder, filename)
    with open(path, "w", encoding="cp950", errors="strict") as f:
        f.write(content)
    print(f"  [OK] {folder}\\{filename}")

###############################################################################
# 1. ChartsNormal - LollipopChartExample
###############################################################################
write_bas("ChartsNormal", "LollipopChartExample.bas", r"""Attribute VB_Name = "LollipopChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: LollipopChartExample
'功能說明: 在Excel中建立棒棒糖圖表（Lollipop Chart）的示範程式，以散佈圖加誤差線呈現
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestLollipopChart()
    Call CreateLollipopChart("棒棒糖圖表")
End Sub

' 建立棒棒糖圖表
' sheetName: 要建立圖表的工作表名稱
Sub CreateLollipopChart(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim dataRange As Range
    Dim ser As Series
    
    ' 取得或建立工作表
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillLollipopData(ws)
    
    ' 先建立散佈圖
    Set dataRange = ws.Range("B2:C8")
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("E1").Left, _
        Top:=ws.Range("E1").Top, _
        Width:=480, _
        Height:=320)
    
    Set cht = chartObj.Chart
    cht.ChartType = xlXYScatter
    cht.SetSourceData Source:=dataRange
    
    ' 設定X軸標籤
    cht.SeriesCollection(1).XValues = ws.Range("C2:C8")
    cht.SeriesCollection(1).Values = ws.Range("B2:B8")
    
    ' 加入垂直誤差線模擬棒棒糖的線條
    Set ser = cht.SeriesCollection(1)
    ser.HasErrorBars = True
    With ser.ErrorBars(xlY)
        .EndStyle = xlNoCap
        .Direction = xlY
        .Include = xlMinusValues
        .Type = xlFixedValue
        .Value = cht.Axes(xlValue).MaximumScale
    End With
    
    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "產品銷售棒棒糖圖"
    
    ' 設定軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "銷售額"
    End With
    
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "產品名稱"
    End With
    
    ' 設定樣式
    cht.ChartStyle = 2
    ser.MarkerSize = 10
    ser.MarkerStyle = xlMarkerStyleCircle
    
    MsgBox "棒棒糖圖表已建立完成！", vbInformation, "完成"
End Sub

' 填入棒棒糖圖表示範資料
Private Sub FillLollipopData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "銷售額"
    ws.Range("C1").Value = "X軸位置"
    
    ws.Range("A2").Value = "產品A"
    ws.Range("B2").Value = 850
    ws.Range("C2").Value = 1
    
    ws.Range("A3").Value = "產品B"
    ws.Range("B3").Value = 620
    ws.Range("C3").Value = 2
    
    ws.Range("A4").Value = "產品C"
    ws.Range("B4").Value = 430
    ws.Range("C4").Value = 3
    
    ws.Range("A5").Value = "產品D"
    ws.Range("B5").Value = 780
    ws.Range("C5").Value = 4
    
    ws.Range("A6").Value = "產品E"
    ws.Range("B6").Value = 550
    ws.Range("C6").Value = 5
    
    ws.Range("A7").Value = "產品F"
    ws.Range("B7").Value = 920
    ws.Range("C7").Value = 6
    
    ws.Range("A8").Value = "產品G"
    ws.Range("B8").Value = 680
    ws.Range("C8").Value = 7
    
    ws.Columns("A:C").AutoFit
End Sub
""")

###############################################################################
# 2. FormulaCreate - ForecastFormulaExample
###############################################################################
write_bas("FormulaCreate", "ForecastFormulaExample.bas", r"""Attribute VB_Name = "ForecastFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: ForecastFormulaExample
'功能說明: 示範透過VBA在Excel中輸入預測公式（FORECAST.LINEAR、TREND、GROWTH）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestForecastFormula()
    Call CreateForecastFormulaExample
End Sub

' 建立預測公式範例
Sub CreateForecastFormulaExample()
    Dim ws As Worksheet
    Dim sheetName As String
    
    sheetName = "預測公式範例"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillForecastData(ws)
    Call EnterForecastFormulas(ws)
    
    ws.Activate
    MsgBox "預測公式範例已建立完成！", vbInformation, "完成"
End Sub

' 填入預測歷史資料
Private Sub FillForecastData(ByVal ws As Worksheet)
    Dim i As Long
    
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "銷售額"
    
    ' 使用迴圈填入資料
    For i = 1 To 12
        ws.Cells(i + 1, 1).Value = "2024/" & i
        ws.Cells(i + 1, 2).Value = 1000 + i * 100 + i * 10
    Next i
End Sub

' 輸入預測公式
Private Sub EnterForecastFormulas(ByVal ws As Worksheet)
    ' 預測目標區
    ws.Range("D1").Value = "預測目標月份"
    ws.Range("E1").Value = "FORECAST.LINEAR"
    ws.Range("F1").Value = "TREND多點預測"
    ws.Range("G1").Value = "GROWTH指數預測"
    
    ws.Range("D2").Value = "2025/1"
    ws.Range("D3").Value = "2025/2"
    ws.Range("D4").Value = "2025/3"
    
    ' FORECAST.LINEAR 線性預測 (Excel 2016+)
    ws.Range("E2").Formula = "=FORECAST.LINEAR(D2,B2:B13,A2:A13)"
    ws.Range("E3").Formula = "=FORECAST.LINEAR(D3,B2:B13,A2:A13)"
    ws.Range("E4").Formula = "=FORECAST.LINEAR(D4,B2:B13,A2:A13)"
    
    ' TREND 函數預測多點
    ws.Range("F2").Formula = "=TREND(B2:B13,A2:A13,D2)"
    ws.Range("F3").Formula = "=TREND(B2:B13,A2:A13,D3)"
    ws.Range("F4").Formula = "=TREND(B2:B13,A2:A13,D4)"
    
    ' GROWTH 指數成長預測
    ws.Range("G2").Formula = "=GROWTH(B2:B13,A2:A13,D2)"
    ws.Range("G3").Formula = "=GROWTH(B2:B13,A2:A13,D3)"
    ws.Range("G4").Formula = "=GROWTH(B2:B13,A2:A13,D4)"
    
    ws.Columns("A:G").AutoFit
End Sub
""")

###############################################################################
# 3. FileMerge - MergeExcelWithFooterRow
###############################################################################
write_bas("FileMerge", "MergeExcelWithFooterRow.bas", r"""Attribute VB_Name = "MergeExcelWithFooterRow"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelWithFooterRow
'功能說明: 合併多個Excel檔案，並在每個來源資料區塊後方自動加上統計頁尾列的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestMergeExcelWithFooterRow()
    Dim folderPath As String
    folderPath = "C:\Temp\MergeData"
    Call MergeExcelWithFooterRow(folderPath)
End Sub

' 合併指定資料夾內所有Excel檔案，並為每個檔案的資料區塊加入小計頁尾列
' folderPath: 來源Excel檔案所在資料夾路徑
Sub MergeExcelWithFooterRow(ByVal folderPath As String)
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim fileName As String
    Dim lastRow As Long
    Dim sourceLastRow As Long
    Dim footerRow As Long
    Dim sumRange As Range
    Dim fso As Object
    Dim folder As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        MsgBox "資料夾不存在：" & folderPath, vbExclamation, "錯誤"
        Set fso = Nothing
        Exit Sub
    End If
    
    ' 建立目標活頁簿
    Set targetWb = Workbooks.Add
    Set targetWs = targetWb.Worksheets(1)
    targetWs.Name = "合併結果"
    
    ' 標題列
    targetWs.Range("A1").Value = "產品"
    targetWs.Range("B1").Value = "數量"
    targetWs.Range("C1").Value = "金額"
    targetWs.Range("D1").Value = "來源檔案"
    
    lastRow = 1
    
    ' 使用 FileSystemObject 遍歷檔案
    Set fso = New Scripting.FileSystemObject
    Dim folderObj As Object
    Dim fileObj As Object
    
    Set folderObj = fso.GetFolder(folderPath)
    
    For Each fileObj In folderObj.Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "xlsx" Then
            fileName = fileObj.Name
            
            Set sourceWb = Workbooks.Open(fileObj.Path)
            Set sourceWs = sourceWb.Worksheets(1)
            sourceLastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
            
            If sourceLastRow > 1 Then
                ' 複製資料（略過標題列）
                sourceWs.Range("A2:C" & sourceLastRow).Copy
                targetWs.Cells(lastRow + 1, 1).PasteSpecial xlPasteValues
                
                ' 填入來源檔案名稱
                targetWs.Range("D" & lastRow + 1 & ":D" & lastRow + sourceLastRow - 1).Value = fileName
                
                ' 計算新的最後列
                lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
                
                ' 加入頁尾列 - 小計
                lastRow = lastRow + 1
                targetWs.Cells(lastRow, 1).Value = "小計 (" & fileName & ")"
                targetWs.Cells(lastRow, 1).Font.Bold = True
                
                ' 數量合計
                footerRow = lastRow
                If sourceLastRow - 1 > 0 Then
                    Set sumRange = targetWs.Range("B" & footerRow - sourceLastRow + 1 & ":B" & footerRow - 1)
                    targetWs.Cells(footerRow, 2).Formula = "=SUM(" & sumRange.Address(False, False) & ")"
                    targetWs.Cells(footerRow, 2).Font.Bold = True
                    
                    Set sumRange = targetWs.Range("C" & footerRow - sourceLastRow + 1 & ":C" & footerRow - 1)
                    targetWs.Cells(footerRow, 3).Formula = "=SUM(" & sumRange.Address(False, False) & ")"
                    targetWs.Cells(footerRow, 3).Font.Bold = True
                End If
                
                ' 加上邊框分隔
                With targetWs.Range("A" & footerRow & ":D" & footerRow).Borders(xlEdgeBottom)
                    .LineStyle = xlDouble
                    .Weight = xlThick
                End With
            End If
            
            sourceWb.Close False
        End If
    Next fileObj
    
    ' 最後加入總計列
    lastRow = targetWs.Cells(targetWs.Rows.Count, 1).End(xlUp).Row
    lastRow = lastRow + 1
    targetWs.Cells(lastRow, 1).Value = "總計"
    targetWs.Cells(lastRow, 1).Font.Bold = True
    targetWs.Cells(lastRow, 2).Formula = "=SUMPRODUCT((LEFT(D2:D" & lastRow - 1 & ",2)<>""小計"")*B2:B" & lastRow - 1 & ")"
    targetWs.Cells(lastRow, 2).Font.Bold = True
    targetWs.Cells(lastRow, 3).Formula = "=SUMPRODUCT((LEFT(D2:D" & lastRow - 1 & ",2)<>""小計"")*C2:C" & lastRow - 1 & ")"
    targetWs.Cells(lastRow, 3).Font.Bold = True
    
    targetWs.Columns("A:D").AutoFit
    
    Set fso = Nothing
    MsgBox "檔案合併完成，已加入頁尾小計列！", vbInformation, "完成"
End Sub
""")

###############################################################################
# 4. FileSplit - SplitByCustomFormula
###############################################################################
write_bas("FileSplit", "SplitByCustomFormula.bas", r"""Attribute VB_Name = "SplitByCustomFormula"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByCustomFormula
'功能說明: 根據使用者自訂公式的計算結果將資料分割到不同工作表的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestSplitByCustomFormula()
    Call SplitByCustomFormula
End Sub

' 根據自訂條件分割資料
Sub SplitByCustomFormula()
    Dim wsSource As Worksheet
    Dim wsPass As Worksheet
    Dim wsFail As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim passRow As Long
    Dim failRow As Long
    Dim score As Double
    
    sheetName = "分割來源"
    
    ' 取得或建立來源工作表
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        Set wsSource = ThisWorkbook.Worksheets.Add
        wsSource.Name = sheetName
    End If
    
    wsSource.Cells.Clear
    Call FillSplitSourceData(wsSource)
    
    ' 建立目標工作表
    On Error Resume Next
    ThisWorkbook.Worksheets("通過").Delete
    ThisWorkbook.Worksheets("未通過").Delete
    On Error GoTo 0
    
    Set wsPass = ThisWorkbook.Worksheets.Add
    wsPass.Name = "通過"
    Set wsFail = ThisWorkbook.Worksheets.Add
    wsFail.Name = "未通過"
    
    ' 複製標題列
    wsSource.Rows(1).Copy wsPass.Rows(1)
    wsSource.Rows(1).Copy wsFail.Rows(1)
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    passRow = 1
    failRow = 1
    
    For i = 2 To lastRow
        score = wsSource.Cells(i, 2).Value
        ' 自訂條件：分數大於等於60為通過
        If score >= 60 Then
            passRow = passRow + 1
            wsSource.Rows(i).Copy wsPass.Rows(passRow)
        Else
            failRow = failRow + 1
            wsSource.Rows(i).Copy wsFail.Rows(failRow)
        End If
    Next i
    
    wsPass.Columns("A:B").AutoFit
    wsFail.Columns("A:B").AutoFit
    
    MsgBox "資料已依自訂公式分割完成！" & vbCrLf & _
           "通過: " & passRow - 1 & " 筆" & vbCrLf & _
           "未通過: " & failRow - 1 & " 筆", vbInformation, "完成"
End Sub

' 填入分割來源資料
Private Sub FillSplitSourceData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "學生姓名"
    ws.Range("B1").Value = "成績"
    
    ws.Range("A2").Value = "王小明"
    ws.Range("B2").Value = 85
    
    ws.Range("A3").Value = "李小華"
    ws.Range("B3").Value = 55
    
    ws.Range("A4").Value = "張大為"
    ws.Range("B4").Value = 72
    
    ws.Range("A5").Value = "陳美玲"
    ws.Range("B5").Value = 90
    
    ws.Range("A6").Value = "林志明"
    ws.Range("B6").Value = 45
    
    ws.Range("A7").Value = "周文彬"
    ws.Range("B7").Value = 68
    
    ws.Range("A8").Value = "吳雅婷"
    ws.Range("B8").Value = 58
    
    ws.Range("A9").Value = "劉建國"
    ws.Range("B9").Value = 77
    
    ws.Columns("A:B").AutoFit
End Sub
""")

###############################################################################
# 5. MergeDataAcrossSheets - MergeWithFilterAndSort
###############################################################################
write_bas("MergeDataAcrossSheets", "MergeWithFilterAndSort.bas", r"""Attribute VB_Name = "MergeWithFilterAndSort"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithFilterAndSort
'功能說明: 合併跨工作表資料，並自動篩選非空白及排序的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestMergeWithFilterAndSort()
    Call MergeWithFilterAndSort
End Sub

' 合併所有工作表資料，篩選指定欄位非空白，並依金額欄排序
Sub MergeWithFilterAndSort()
    Dim wsTarget As Worksheet
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim sourceLastRow As Long
    Dim i As Long
    
    ' 建立目標工作表
    On Error Resume Next
    ThisWorkbook.Worksheets("合併篩選排序").Delete
    On Error GoTo 0
    
    Set wsTarget = ThisWorkbook.Worksheets.Add
    wsTarget.Name = "合併篩選排序"
    wsTarget.Range("A1").Value = "產品名稱"
    wsTarget.Range("B1").Value = "分類"
    wsTarget.Range("C1").Value = "金額"
    wsTarget.Range("D1").Value = "來源工作表"
    
    targetRow = 1
    
    ' 遍歷所有工作表
    For Each ws In ThisWorkbook.Worksheets
        ' 跳過目標工作表
        If ws.Name <> "合併篩選排序" Then
            sourceLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            If sourceLastRow > 1 Then
                ' 複製資料（略過標題列）
                ws.Range("A2:C" & sourceLastRow).Copy
                wsTarget.Cells(targetRow + 1, 1).PasteSpecial xlPasteValues
                
                ' 填入來源工作表名稱
                wsTarget.Range("D" & targetRow + 1 & ":D" & targetRow + sourceLastRow - 1).Value = ws.Name
            End If
            
            targetRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
        End If
    Next ws
    
    ' 刪除C欄為空白的列（篩選功能）
    If targetRow > 1 Then
        For i = targetRow To 2 Step -1
            If IsEmpty(wsTarget.Cells(i, 3)) Or Len(CStr(wsTarget.Cells(i, 3).Value)) = 0 Then
                wsTarget.Rows(i).Delete
            End If
        Next i
    End If
    
    ' 重新取得最後列
    targetRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
    
    ' 依金額欄排序（降冪）
    If targetRow > 1 Then
        With wsTarget.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsTarget.Range("C2:C" & targetRow), _
                SortOn:=xlSortOnValues, _
                Order:=xlDescending
            .SetRange wsTarget.Range("A1:D" & targetRow)
            .Header = xlYes
            .Apply
        End With
    End If
    
    wsTarget.Columns("A:D").AutoFit
    wsTarget.Activate
    
    MsgBox "跨工作表資料已合併、篩選並排序完成！", vbInformation, "完成"
End Sub
""")

###############################################################################
# 6. ExporttoPDF - ExportPDFByCellValueCondition
###############################################################################
write_bas("ExporttoPDF", "ExportPDFByCellValueCondition.bas", r"""Attribute VB_Name = "ExportPDFByCellValueCondition"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFByCellValueCondition
'功能說明: 根據儲存格值條件判斷是否匯出PDF，並以儲存格內容命名檔案的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestExportPDFByCellValueCondition()
    Call ExportPDFByCellValueCondition("D:\Temp\PDFOutput")
End Sub

' 根據儲存格值條件匯出PDF
' outputPath: PDF輸出資料夾路徑
Sub ExportPDFByCellValueCondition(ByVal outputPath As String)
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim exportCount As Long
    Dim fileName As String
    Dim status As String
    Dim pdfPath As String
    Dim fso As Object
    Dim invalidChars As Variant
    Dim j As Long
    
    sheetName = "PDF條件匯出"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillPDFExportData(ws)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(outputPath) Then
        fso.CreateFolder outputPath
    End If
    
    exportCount = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 無效檔案名字元陣列
    invalidChars = Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
    
    For i = 2 To lastRow
        status = CStr(ws.Cells(i, 4).Value)
        
        ' 條件：只有狀態為"已完成"的才匯出
        If status = "已完成" Then
            fileName = CStr(ws.Cells(i, 1).Value) & "_" & _
                       CStr(ws.Cells(i, 2).Value) & ".pdf"
            
            ' 將無效檔名字元替換
            For j = LBound(invalidChars) To UBound(invalidChars)
                fileName = Replace(fileName, CStr(invalidChars(j)), "_")
            Next j
            
            pdfPath = outputPath & "\" & fileName
            
            ' 將該列資料標記到暫時區域並匯出
            ws.Cells(i, 1).Resize(1, 4).Copy
            ws.Range("F1").PasteSpecial xlPasteValues
            
            ' 設定列印範圍並匯出PDF
            ws.PageSetup.PrintArea = "F1:I1"
            ws.ExportAsFixedFormat Type:=xlTypePDF, _
                fileName:=pdfPath, _
                Quality:=xlQualityStandard
            
            exportCount = exportCount + 1
        End If
    Next i
    
    ws.Range("F1:I1").Clear
    
    Set fso = Nothing
    
    MsgBox "條件式PDF匯出完成！" & vbCrLf & _
           "共匯出 " & exportCount & " 個PDF檔案至：" & vbCrLf & _
           outputPath, vbInformation, "完成"
End Sub

' 填入PDF條件匯出示範資料
Private Sub FillPDFExportData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "客戶編號"
    ws.Range("B1").Value = "客戶名稱"
    ws.Range("C1").Value = "金額"
    ws.Range("D1").Value = "狀態"
    
    ws.Range("A2").Value = "C001"
    ws.Range("B2").Value = "大華貿易"
    ws.Range("C2").Value = 15000
    ws.Range("D2").Value = "已完成"
    
    ws.Range("A3").Value = "C002"
    ws.Range("B3").Value = "明遠科技"
    ws.Range("C3").Value = 8500
    ws.Range("D3").Value = "進行中"
    
    ws.Range("A4").Value = "C003"
    ws.Range("B4").Value = "金茂實業"
    ws.Range("C4").Value = 22000
    ws.Range("D4").Value = "已完成"
    
    ws.Range("A5").Value = "C004"
    ws.Range("B5").Value = "永豐物流"
    ws.Range("C5").Value = 12000
    ws.Range("D5").Value = "已完成"
    
    ws.Columns("A:D").AutoFit
End Sub
""")

###############################################################################
# 7. ConditionalFormatting - RowColumnCrossHighlight
###############################################################################
write_bas("ConditionalFormatting", "RowColumnCrossHighlight.bas", r"""Attribute VB_Name = "RowColumnCrossHighlight"
Option Explicit
'*************************************************************************************
'模組名稱: RowColumnCrossHighlight
'功能說明: 使用VBA搭配條件式格式，實現選取儲存格時自動交叉亮顯該列與該欄的效果
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 建立交叉亮顯範例工作表
Sub CreateCrossHighlightExample()
    Dim ws As Worksheet
    Dim sheetName As String
    
    sheetName = "交叉亮顯"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillCrossHighlightData(ws)
    Call SetupCrossHighlight(ws)
    
    ws.Activate
    MsgBox "交叉亮顯已設定完成！" & vbCrLf & _
           "請嘗試點選任一儲存格查看效果。", vbInformation, "完成"
End Sub

' 設定交叉亮顯條件式格式
Private Sub SetupCrossHighlight(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim lastCol As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    With ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .Interior.ColorIndex = xlNone
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    
    ' 提示使用者可在對應工作表的SelectionChange事件中使用本模組的HighlightCrossRowColumn程序
    MsgBox "請將本模組中的 HighlightCrossRowColumn 程序" & vbCrLf & _
           "複製到對應工作表的 Worksheet_SelectionChange 事件中即可。", _
           vbInformation, "提示"
End Sub

' 交叉亮顯處理程序（可放入Worksheet_SelectionChange事件中）
Public Sub HighlightCrossRowColumn(ByVal Target As Range)
    Dim ws As Worksheet
    Dim rng As Range
    
    Set ws = Target.Worksheet
    
    ' 清除所有儲存格的條件式格式
    ws.Cells.FormatConditions.Delete
    
    ' 設定選取列的條件式格式（淡黃色）
    With ws.Rows(Target.Row)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=ROW()=" & Target.Row
        .FormatConditions(1).Interior.Color = RGB(255, 255, 200)
    End With
    
    ' 設定選取欄的條件式格式（淡綠色）
    With ws.Columns(Target.Column)
        .FormatConditions.Add Type:=xlExpression, Formula1:="=COLUMN()=" & Target.Column
        .FormatConditions(1).Interior.Color = RGB(200, 255, 200)
    End With
    
    ' 選取儲存格交叉處特別標示
    With Target
        .FormatConditions.Add Type:=xlExpression, _
            Formula1:="=AND(ROW()=" & Target.Row & ",COLUMN()=" & Target.Column & ")"
        .FormatConditions(1).Interior.Color = RGB(255, 150, 100)
    End With
End Sub

' 填入交叉亮顯示範資料
Private Sub FillCrossHighlightData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "品名"
    ws.Range("B1").Value = "一月"
    ws.Range("C1").Value = "二月"
    ws.Range("D1").Value = "三月"
    ws.Range("E1").Value = "四月"
    ws.Range("F1").Value = "合計"
    
    ws.Range("A2").Value = "產品A"
    ws.Range("B2").Value = 100
    ws.Range("C2").Value = 120
    ws.Range("D2").Value = 140
    ws.Range("E2").Value = 160
    ws.Range("F2").Formula = "=SUM(B2:E2)"
    
    ws.Range("A3").Value = "產品B"
    ws.Range("B3").Value = 80
    ws.Range("C3").Value = 85
    ws.Range("D3").Value = 90
    ws.Range("E3").Value = 95
    ws.Range("F3").Formula = "=SUM(B3:E3)"
    
    ws.Range("A4").Value = "產品C"
    ws.Range("B4").Value = 200
    ws.Range("C4").Value = 180
    ws.Range("D4").Value = 160
    ws.Range("E4").Value = 140
    ws.Range("F4").Formula = "=SUM(B4:E4)"
    
    ws.Range("A5").Value = "產品D"
    ws.Range("B5").Value = 150
    ws.Range("C5").Value = 170
    ws.Range("D5").Value = 190
    ws.Range("E5").Value = 210
    ws.Range("F5").Formula = "=SUM(B5:E5)"
    
    ws.Columns("A:F").AutoFit
End Sub
""")

###############################################################################
# 8. ClearCellFormatting - ClearNonStandardFormatting
###############################################################################
write_bas("ClearCellFormatting", "ClearNonStandardFormatting.bas", r"""Attribute VB_Name = "ClearNonStandardFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearNonStandardFormatting
'功能說明: 清除非標準格式的儲存格，僅保留指定的字型、大小、顏色、框線等標準格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestClearNonStandardFormatting()
    Call ClearNonStandardFormatting
End Sub

' 清除非標準格式，只保留指定格式
Sub ClearNonStandardFormatting()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim targetRange As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim lastCol As Long
    
    sheetName = "清除非標準格式"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillClearFormatData(ws)
    
    ' 顯示清除前的狀態
    MsgBox "即將清除非標準格式。" & vbCrLf & _
           "標準格式定義：字型=微軟正黑體, 大小=11, 字型色=黑色", vbInformation, "提示"
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    For Each cell In targetRange
        With cell
            ' 清除非標準字型
            If .Font.Name <> "微軟正黑體" Then
                .Font.Name = "微軟正黑體"
            End If
            
            ' 清除非標準字型大小
            If .Font.Size <> 11 Then
                .Font.Size = 11
            End If
            
            ' 清除非標準字型色（非黑色）
            If .Font.Color <> RGB(0, 0, 0) Then
                .Font.Color = RGB(0, 0, 0)
            End If
            
            ' 清除非標準粗體
            If .Font.Bold = True Then
                .Font.Bold = False
            End If
            
            ' 清除非標準斜體
            If .Font.Italic = True Then
                .Font.Italic = False
            End If
            
            ' 清除非標準底線
            If .Font.Underline <> xlUnderlineStyleNone Then
                .Font.Underline = xlUnderlineStyleNone
            End If
            
            ' 清除非白色/無填滿的背景色
            If .Interior.Color <> RGB(255, 255, 255) And _
               .Interior.ColorIndex <> xlNone Then
                .Interior.ColorIndex = xlNone
            End If
            
            ' 清除內部框線（只保留最外框）
            If .Borders(xlInsideVertical).LineStyle <> xlNone Then
                .Borders(xlInsideVertical).LineStyle = xlNone
            End If
            If .Borders(xlInsideHorizontal).LineStyle <> xlNone Then
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            End If
        End With
    Next cell
    
    ws.Columns("A:C").AutoFit
    ws.Activate
    
    MsgBox "非標準格式已清除完成！", vbInformation, "完成"
End Sub

' 填入示範資料（含多種不同格式）
Private Sub FillClearFormatData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品名稱"
    ws.Range("B1").Value = "數量"
    ws.Range("C1").Value = "備註"
    
    ' 標題列使用特殊格式
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(0, 0, 255)
        .Interior.Color = RGB(200, 200, 255)
    End With
    
    ws.Range("A2").Value = "產品X"
    ws.Range("B2").Value = 100
    ws.Range("C2").Value = "正常"
    ws.Range("A2").Font.Italic = True
    
    ws.Range("A3").Value = "產品Y"
    ws.Range("B3").Value = 200
    ws.Range("C3").Value = "測試"
    ws.Range("A3").Font.Underline = xlUnderlineStyleSingle
    
    ws.Range("A4").Value = "產品Z"
    ws.Range("B4").Value = 300
    ws.Range("C4").Value = "特例"
    ws.Range("A4").Interior.Color = RGB(255, 255, 0)
    
    ws.Range("A5").Value = "產品W"
    ws.Range("B5").Value = 400
    ws.Range("C5").Value = "緊急"
    ws.Range("A5").Font.Color = RGB(255, 0, 0)
    ws.Range("A5").Font.Size = 16
    
    ws.Columns("A:C").AutoFit
End Sub
""")

###############################################################################
# 9. BatchEnterFormulas - BatchDATEDIFFormulas
###############################################################################
write_bas("BatchEnterFormulas", "BatchDATEDIFFormulas.bas", r"""Attribute VB_Name = "BatchDATEDIFFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchDATEDIFFormulas
'功能說明: 批次輸入DATEDIF日期差異公式（年、月、日差異計算）的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestBatchDATEDIFFormulas()
    Call BatchEnterDATEDIFFormulas
End Sub

' 批次輸入DATEDIF公式
Sub BatchEnterDATEDIFFormulas()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    
    sheetName = "DATEDIF公式"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillDATEDIFData(ws)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 批次輸入DATEDIF公式
    For i = 2 To lastRow
        ' 計算年數差異
        ws.Cells(i, 4).Formula = "=DATEDIF(A" & i & ",B" & i & ",""Y"")"
        
        ' 計算總月數差異
        ws.Cells(i, 5).Formula = "=DATEDIF(A" & i & ",B" & i & ",""M"")"
        
        ' 計算月數（去除整年後剩餘月數）
        ws.Cells(i, 6).Formula = "=DATEDIF(A" & i & ",B" & i & ",""YM"")"
        
        ' 計算總天數差異
        ws.Cells(i, 7).Formula = "=DATEDIF(A" & i & ",B" & i & ",""D"")"
        
        ' 計算天數（去除整月後剩餘天數）
        ws.Cells(i, 8).Formula = "=DATEDIF(A" & i & ",B" & i & ",""MD"")"
        
        ' 組合文字格式的年月日差異
        ws.Cells(i, 9).Formula = _
            "=DATEDIF(A" & i & ",B" & i & ",""Y"")&""年""&" & _
            "DATEDIF(A" & i & ",B" & i & ",""YM"")&""個月""&" & _
            "DATEDIF(A" & i & ",B" & i & ",""MD"")&" & Chr(34) & "天" & Chr(34)
    Next i
    
    ws.Columns("A:I").AutoFit
    ws.Activate
    
    MsgBox "DATEDIF公式批次輸入完成！共 " & lastRow - 1 & " 筆。", vbInformation, "完成"
End Sub

' 填入DATEDIF示範資料
Private Sub FillDATEDIFData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "開始日期"
    ws.Range("B1").Value = "結束日期"
    ws.Range("C1").Value = "專案名稱"
    ws.Range("D1").Value = "年數(Y)"
    ws.Range("E1").Value = "總月數(M)"
    ws.Range("F1").Value = "剩餘月數(YM)"
    ws.Range("G1").Value = "總天數(D)"
    ws.Range("H1").Value = "剩餘天數(MD)"
    ws.Range("I1").Value = "年月日差異"
    
    ws.Range("A2").Value = "2020/1/15"
    ws.Range("B2").Value = "2024/6/1"
    ws.Range("C2").Value = "專案A"
    
    ws.Range("A3").Value = "2018/3/10"
    ws.Range("B3").Value = "2023/12/31"
    ws.Range("C3").Value = "專案B"
    
    ws.Range("A4").Value = "2021/7/1"
    ws.Range("B4").Value = "2025/3/15"
    ws.Range("C4").Value = "專案C"
    
    ws.Range("A5").Value = "2019/11/20"
    ws.Range("B5").Value = "2024/8/25"
    ws.Range("C5").Value = "專案D"
    
    ws.Range("A6").Value = "2022/1/5"
    ws.Range("B6").Value = "2025/1/5"
    ws.Range("C6").Value = "專案E"
    
    ws.Range("A7").Value = "2020/6/15"
    ws.Range("B7").Value = "2024/6/15"
    ws.Range("C7").Value = "專案F"
    
    ws.Columns("A:C").AutoFit
End Sub
""")

###############################################################################
# 10. AutomaticallyCompareDataDifferences - CompareWithChartVisual
###############################################################################
write_bas("AutomaticallyCompareDataDifferences", "CompareWithChartVisual.bas", r"""Attribute VB_Name = "CompareWithChartVisual"
Option Explicit
'*************************************************************************************
'模組名稱: CompareWithChartVisual
'功能說明: 自動比較兩組資料差異，並以圖表視覺化呈現比較結果的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestCompareWithChartVisual()
    Call CompareWithChartVisual
End Sub

' 比較兩組資料並以圖表呈現差異
Sub CompareWithChartVisual()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim lastRow As Long
    Dim i As Long
    
    sheetName = "圖表比較差異"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillCompareData(ws)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 計算差異值與差異百分比
    ws.Range("D1").Value = "差異值"
    ws.Range("E1").Value = "差異百分比"
    
    For i = 2 To lastRow
        ws.Cells(i, 4).Formula = "=C" & i & "-B" & i
        ws.Cells(i, 5).Formula = "=IF(B" & i & "=0,0,(C" & i & "-B" & i & ")/B" & i & ")"
        ws.Cells(i, 5).NumberFormat = "0.0%"
    Next i
    
    ' 建立比較圖表
    Set chartObj = ws.ChartObjects.Add( _
        Left:=ws.Range("G1").Left, _
        Top:=ws.Range("G1").Top, _
        Width:=500, _
        Height:=320)
    
    Set cht = chartObj.Chart
    cht.ChartType = xlColumnClustered
    
    ' 加入去年資料系列
    cht.SeriesCollection.NewSeries
    cht.SeriesCollection(1).Name = "去年"
    cht.SeriesCollection(1).Values = ws.Range("B2:B" & lastRow)
    cht.SeriesCollection(1).XValues = ws.Range("A2:A" & lastRow)
    
    ' 加入今年資料系列
    cht.SeriesCollection.NewSeries
    cht.SeriesCollection(2).Name = "今年"
    cht.SeriesCollection(2).Values = ws.Range("C2:C" & lastRow)
    cht.SeriesCollection(2).XValues = ws.Range("A2:A" & lastRow)
    
    ' 加入差異值折線
    cht.SeriesCollection.NewSeries
    cht.SeriesCollection(3).Name = "差異值"
    cht.SeriesCollection(3).Values = ws.Range("D2:D" & lastRow)
    cht.SeriesCollection(3).ChartType = xlLineMarkers
    cht.SeriesCollection(3).XValues = ws.Range("A2:A" & lastRow)
    
    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "年度銷售比較與差異分析"
    
    ' 設定軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "月份"
    End With
    
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "銷售額"
    End With
    
    ' 設定圖例
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom
    
    ' 樣式設定
    cht.ChartStyle = 10
    
    ' 用條件式格式標示差異
    With ws.Range("D2:D" & lastRow)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Interior.Color = RGB(200, 255, 200)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(2).Interior.Color = RGB(255, 200, 200)
    End With
    
    ws.Columns("A:E").AutoFit
    ws.Activate
    
    MsgBox "資料差異比較與圖表視覺化已完成！", vbInformation, "完成"
End Sub

' 填入比較示範資料
Private Sub FillCompareData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "月份"
    ws.Range("B1").Value = "去年銷售額"
    ws.Range("C1").Value = "今年銷售額"
    
    ws.Range("A2").Value = "1月"
    ws.Range("B2").Value = 1200
    ws.Range("C2").Value = 1350
    
    ws.Range("A3").Value = "2月"
    ws.Range("B3").Value = 1100
    ws.Range("C3").Value = 1080
    
    ws.Range("A4").Value = "3月"
    ws.Range("B4").Value = 1400
    ws.Range("C4").Value = 1550
    
    ws.Range("A5").Value = "4月"
    ws.Range("B5").Value = 1300
    ws.Range("C5").Value = 1420
    
    ws.Range("A6").Value = "5月"
    ws.Range("B6").Value = 1500
    ws.Range("C6").Value = 1600
    
    ws.Range("A7").Value = "6月"
    ws.Range("B7").Value = 1600
    ws.Range("C7").Value = 1480
    
    ws.Columns("A:C").AutoFit
End Sub
""")

###############################################################################
# 11. AutomaticallyCleanData - CleanDateTimeSeparate
###############################################################################
write_bas("AutomaticallyCleanData", "CleanDateTimeSeparate.bas", r"""Attribute VB_Name = "CleanDateTimeSeparate"
Option Explicit
'*************************************************************************************
'模組名稱: CleanDateTimeSeparate
'功能說明: 自動清理日期時間資料，將混合的日期時間欄位拆分為獨立的日期欄與時間欄
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestCleanDateTimeSeparate()
    Call CleanDateTimeSeparate
End Sub

' 清理並分離日期時間資料
Sub CleanDateTimeSeparate()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim rawValue As String
    Dim dateValue As Date
    Dim datePart As String
    Dim timePart As String
    
    sheetName = "日期時間清理"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillDateTimeData(ws)
    
    ' 標題列
    ws.Range("D1").Value = "日期"
    ws.Range("E1").Value = "時間"
    ws.Range("F1").Value = "清理狀態"
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        rawValue = CStr(ws.Cells(i, 3).Value)
        
        On Error Resume Next
        ' 嘗試轉換為日期
        dateValue = CDate(rawValue)
        
        If Err.Number = 0 Then
            ' 成功轉換，分離日期和時間
            datePart = Format(dateValue, "yyyy/mm/dd")
            timePart = Format(dateValue, "hh:mm:ss")
            
            ws.Cells(i, 4).Value = datePart
            ws.Cells(i, 5).Value = timePart
            
            ' 設定日期格式
            ws.Cells(i, 4).NumberFormat = "yyyy/mm/dd"
            ws.Cells(i, 5).NumberFormat = "hh:mm:ss"
            
            ws.Cells(i, 6).Value = "已清理"
            ws.Cells(i, 6).Interior.Color = RGB(200, 255, 200)
        Else
            ' 失敗，標記為異常
            ws.Cells(i, 4).Value = "格式錯誤"
            ws.Cells(i, 5).Value = "格式錯誤"
            ws.Cells(i, 6).Value = "異常"
            ws.Cells(i, 6).Interior.Color = RGB(255, 200, 200)
            Err.Clear
        End If
        On Error GoTo 0
    Next i
    
    ws.Columns("A:F").AutoFit
    ws.Activate
    
    MsgBox "日期時間資料清理與分離完成！共處理 " & lastRow - 1 & " 筆資料。", vbInformation, "完成"
End Sub

' 填入日期時間示範資料（含多種格式）
Private Sub FillDateTimeData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "記錄編號"
    ws.Range("B1").Value = "事件名稱"
    ws.Range("C1").Value = "原始日期時間"
    
    ws.Range("A2").Value = 1
    ws.Range("B2").Value = "系統登入"
    ws.Range("C2").Value = "2024/6/15 08:30:00"
    
    ws.Range("A3").Value = 2
    ws.Range("B3").Value = "訂單建立"
    ws.Range("C3").Value = "2024-06-15 14:25:30"
    
    ws.Range("A4").Value = 3
    ws.Range("B4").Value = "出貨確認"
    ws.Range("C4").Value = "2024/6/16 10:15"
    
    ws.Range("A5").Value = 4
    ws.Range("B5").Value = "付款通知"
    ws.Range("C5").Value = "2024/06/17 16:45:22"
    
    ws.Range("A6").Value = 5
    ws.Range("B6").Value = "退貨處理"
    ws.Range("C6").Value = "N/A"
    
    ws.Range("A7").Value = 6
    ws.Range("B7").Value = "庫存更新"
    ws.Range("C7").Value = "2024/6/18 09:00:00"
    
    ws.Range("A8").Value = 7
    ws.Range("B8").Value = "客戶回覆"
    ws.Range("C8").Value = "無時間記錄"
    
    ws.Columns("A:C").AutoFit
End Sub
""")

###############################################################################
# 12. FilterDataBasedonMultipleConditions - FilterByTimeRange
###############################################################################
write_bas("FilterDataBasedonMultipleConditions", "FilterByTimeRange.bas", r"""Attribute VB_Name = "FilterByTimeRange"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByTimeRange
'功能說明: 依據多個時間範圍條件（日期範圍+時段範圍）篩選資料的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestFilterByTimeRange()
    Call FilterByTimeRange(#2024/6/1#, #2024/6/30#, "09:00", "17:00")
End Sub

' 依時間範圍篩選資料
' startDate: 開始日期
' endDate: 結束日期
' startTime: 開始時間（如 "09:00"）
' endTime: 結束時間（如 "17:00"）
Sub FilterByTimeRange(ByVal startDate As Date, ByVal endDate As Date, _
                       ByVal startTime As String, ByVal endTime As String)
    Dim wsSource As Worksheet
    Dim wsResult As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim resultRow As Long
    Dim recordDate As Date
    Dim recordTime As Date
    Dim startTimeVal As Date
    Dim endTimeVal As Date
    Dim dateTimeStr As String
    Dim matchCount As Long
    
    sheetName = "時間篩選來源"
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        Set wsSource = ThisWorkbook.Worksheets.Add
        wsSource.Name = sheetName
    End If
    
    wsSource.Cells.Clear
    Call FillTimeRangeData(wsSource)
    
    ' 建立結果工作表
    On Error Resume Next
    ThisWorkbook.Worksheets("時間篩選結果").Delete
    On Error GoTo 0
    
    Set wsResult = ThisWorkbook.Worksheets.Add
    wsResult.Name = "時間篩選結果"
    wsSource.Rows(1).Copy wsResult.Rows(1)
    resultRow = 1
    
    ' 轉換時間字串
    startTimeVal = TimeValue(startTime)
    endTimeVal = TimeValue(endTime)
    
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    matchCount = 0
    
    For i = 2 To lastRow
        dateTimeStr = CStr(wsSource.Cells(i, 3).Value)
        
        On Error Resume Next
        recordDate = CDate(dateTimeStr)
        
        If Err.Number = 0 Then
            recordTime = TimeValue(Format(recordDate, "hh:mm:ss"))
            
            ' 檢查日期範圍與時間範圍
            If recordDate >= startDate And recordDate <= endDate Then
                If recordTime >= startTimeVal And recordTime <= endTimeVal Then
                    resultRow = resultRow + 1
                    wsSource.Rows(i).Copy wsResult.Rows(resultRow)
                    matchCount = matchCount + 1
                End If
            End If
        End If
        Err.Clear
        On Error GoTo 0
    Next i
    
    wsResult.Columns("A:D").AutoFit
    wsResult.Activate
    
    MsgBox "時間範圍篩選完成！" & vbCrLf & _
           "日期範圍：" & startDate & " ~ " & endDate & vbCrLf & _
           "時間範圍：" & startTime & " ~ " & endTime & vbCrLf & _
           "符合筆數：" & matchCount & " 筆", vbInformation, "完成"
End Sub

' 填入時間篩選示範資料
Private Sub FillTimeRangeData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "員工編號"
    ws.Range("B1").Value = "姓名"
    ws.Range("C1").Value = "打卡時間"
    ws.Range("D1").Value = "類型"
    
    ws.Range("A2").Value = "E001"
    ws.Range("B2").Value = "王小明"
    ws.Range("C2").Value = "2024/6/15 08:30:00"
    ws.Range("D2").Value = "上班"
    
    ws.Range("A3").Value = "E002"
    ws.Range("B3").Value = "李小華"
    ws.Range("C3").Value = "2024/6/15 07:45:00"
    ws.Range("D3").Value = "上班"
    
    ws.Range("A4").Value = "E001"
    ws.Range("B4").Value = "王小明"
    ws.Range("C4").Value = "2024/6/15 19:30:00"
    ws.Range("D4").Value = "下班"
    
    ws.Range("A5").Value = "E003"
    ws.Range("B5").Value = "張大為"
    ws.Range("C5").Value = "2024/6/20 09:15:00"
    ws.Range("D5").Value = "上班"
    
    ws.Range("A6").Value = "E002"
    ws.Range("B6").Value = "李小華"
    ws.Range("C6").Value = "2024/6/15 18:00:00"
    ws.Range("D6").Value = "下班"
    
    ws.Range("A7").Value = "E001"
    ws.Range("B7").Value = "王小明"
    ws.Range("C7").Value = "2024/7/1 08:30:00"
    ws.Range("D7").Value = "上班"
    
    ws.Range("A8").Value = "E004"
    ws.Range("B8").Value = "陳美玲"
    ws.Range("C8").Value = "2024/6/25 12:00:00"
    ws.Range("D8").Value = "午休"
    
    ws.Columns("A:D").AutoFit
End Sub
""")

###############################################################################
# 13. PivotTableAnalysis - PivotCacheManagement
###############################################################################
write_bas("PivotTableAnalysis", "PivotCacheManagement.bas", r"""Attribute VB_Name = "PivotCacheManagement"
Option Explicit
'*************************************************************************************
'模組名稱: PivotCacheManagement
'功能說明: 示範樞紐分析表快取管理，包括共用快取、快取資訊查詢與最佳化的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestPivotCacheManagement()
    Call PivotCacheManagementDemo
End Sub

' 樞紐快取管理示範
Sub PivotCacheManagementDemo()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim wsPivot2 As Worksheet
    Dim wsInfo As Worksheet
    Dim dataRange As Range
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim pt2 As PivotTable
    Dim lastRow As Long
    Dim i As Long
    
    ' 建立資料工作表
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("快取管理資料")
    On Error GoTo 0
    
    If wsData Is Nothing Then
        Set wsData = ThisWorkbook.Worksheets.Add
        wsData.Name = "快取管理資料"
    End If
    
    wsData.Cells.Clear
    Call FillCacheDemoData(wsData)
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Set dataRange = wsData.Range("A1:C" & lastRow)
    
    ' 建立第一個樞紐分析表（使用新快取）
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("樞紐共用快取1")
    If Not wsPivot Is Nothing Then wsPivot.Delete
    On Error GoTo 0
    
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "樞紐共用快取1"
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange, _
        Version:=xlPivotTableVersion15)
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A1"), _
        TableName:="PivotShared1")
    
    With pt
        .PivotFields("分類").Orientation = xlRowField
        .PivotFields("分類").Position = 1
        .PivotFields("月份").Orientation = xlColumnField
        .PivotFields("月份").Position = 1
        .PivotFields("金額").Orientation = xlDataField
    End With
    
    ' 建立第二個樞紐分析表（共用相同快取）
    On Error Resume Next
    Set wsPivot2 = ThisWorkbook.Worksheets("樞紐共用快取2")
    If Not wsPivot2 Is Nothing Then wsPivot2.Delete
    On Error GoTo 0
    
    Set wsPivot2 = ThisWorkbook.Worksheets.Add
    wsPivot2.Name = "樞紐共用快取2"
    
    Set pt2 = pc.CreatePivotTable( _
        TableDestination:=wsPivot2.Range("A1"), _
        TableName:="PivotShared2")
    
    With pt2
        .PivotFields("分類").Orientation = xlRowField
        .PivotFields("分類").Position = 1
    End With
    
    pt2.AddDataField pt2.PivotFields("金額"), "金額合計", xlSum
    
    ' 建立快取資訊工作表
    On Error Resume Next
    Set wsInfo = ThisWorkbook.Worksheets("快取資訊")
    If Not wsInfo Is Nothing Then wsInfo.Delete
    On Error GoTo 0
    
    Set wsInfo = ThisWorkbook.Worksheets.Add
    wsInfo.Name = "快取資訊"
    
    wsInfo.Range("A1").Value = "樞紐快取管理資訊"
    wsInfo.Range("A1").Font.Bold = True
    wsInfo.Range("A2").Value = "樞紐快取總數："
    wsInfo.Range("B2").Value = ThisWorkbook.PivotCaches.Count
    
    wsInfo.Range("A4").Value = "快取索引"
    wsInfo.Range("B4").Value = "資料筆數"
    
    For i = 1 To ThisWorkbook.PivotCaches.Count
        wsInfo.Cells(4 + i, 1).Value = ThisWorkbook.PivotCaches(i).Index
        
        On Error Resume Next
        wsInfo.Cells(4 + i, 2).Value = ThisWorkbook.PivotCaches(i).RecordCount
        On Error GoTo 0
    Next i
    
    Dim infoRow As Long
    infoRow = 4 + ThisWorkbook.PivotCaches.Count + 1
    wsInfo.Cells(infoRow, 1).Value = "說明：以上兩個樞紐分析表共用同一個快取，可節省記憶體。"
    
    wsInfo.Columns("A:C").AutoFit
    wsInfo.Activate
    
    MsgBox "樞紐快取管理示範完成！" & vbCrLf & _
           "請查看「快取資訊」工作表了解快取使用狀況。", vbInformation, "完成"
End Sub

' 填入快取管理示範資料
Private Sub FillCacheDemoData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "分類"
    ws.Range("B1").Value = "月份"
    ws.Range("C1").Value = "金額"
    
    ws.Range("A2").Value = "食品"
    ws.Range("B2").Value = "1月"
    ws.Range("C2").Value = 5000
    
    ws.Range("A3").Value = "飲料"
    ws.Range("B3").Value = "1月"
    ws.Range("C3").Value = 3200
    
    ws.Range("A4").Value = "食品"
    ws.Range("B4").Value = "2月"
    ws.Range("C4").Value = 4800
    
    ws.Range("A5").Value = "飲料"
    ws.Range("B5").Value = "2月"
    ws.Range("C5").Value = 3500
    
    ws.Range("A6").Value = "食品"
    ws.Range("B6").Value = "3月"
    ws.Range("C6").Value = 6200
    
    ws.Range("A7").Value = "飲料"
    ws.Range("B7").Value = "3月"
    ws.Range("C7").Value = 4100
    
    ws.Range("A8").Value = "百貨"
    ws.Range("B8").Value = "1月"
    ws.Range("C8").Value = 7200
    
    ws.Range("A9").Value = "百貨"
    ws.Range("B9").Value = "2月"
    ws.Range("C9").Value = 6800
    
    ws.Range("A10").Value = "百貨"
    ws.Range("B10").Value = "3月"
    ws.Range("C10").Value = 8100
    
    ws.Columns("A:C").AutoFit
End Sub
""")

###############################################################################
# 14. PivotCharts - PivotBubbleChartExample
###############################################################################
write_bas("PivotCharts", "PivotBubbleChartExample.bas", r"""Attribute VB_Name = "PivotBubbleChartExample"
Option Explicit
'*************************************************************************************
'模組名稱: PivotBubbleChartExample
'功能說明: 建立樞紐分析表氣泡圖的示範程式，結合樞紐分析與氣泡圖呈現三維度資料
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestPivotBubbleChart()
    Call CreatePivotBubbleChart
End Sub

' 建立樞紐氣泡圖
Sub CreatePivotBubbleChart()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim dataRange As Range
    Dim pc As PivotCache
    Dim pt As PivotTable
    Dim chartObj As ChartObject
    Dim cht As Chart
    Dim lastRow As Long
    
    ' 建立資料工作表
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("氣泡圖資料")
    On Error GoTo 0
    
    If wsData Is Nothing Then
        Set wsData = ThisWorkbook.Worksheets.Add
        wsData.Name = "氣泡圖資料"
    End If
    
    wsData.Cells.Clear
    Call FillBubbleChartData(wsData)
    
    ' 建立樞紐分析表
    On Error Resume Next
    Set wsPivot = ThisWorkbook.Worksheets("氣泡圖樞紐")
    If Not wsPivot Is Nothing Then wsPivot.Delete
    On Error GoTo 0
    
    Set wsPivot = ThisWorkbook.Worksheets.Add
    wsPivot.Name = "氣泡圖樞紐"
    
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    Set dataRange = wsData.Range("A1:D" & lastRow)
    
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=dataRange, _
        Version:=xlPivotTableVersion15)
    
    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A1"), _
        TableName:="PivotBubble")
    
    With pt
        .PivotFields("產品").Orientation = xlRowField
        .PivotFields("產品").Position = 1
    End With
    
    pt.AddDataField pt.PivotFields("銷售額"), "銷售額合計", xlSum
    pt.AddDataField pt.PivotFields("利潤"), "利潤合計", xlSum
    pt.AddDataField pt.PivotFields("市佔率"), "市佔率平均", xlAverage
    
    ' 取得樞紐分析表範圍
    Dim pivotRange As Range
    Dim rowCount As Long
    
    Set pivotRange = pt.TableRange1
    rowCount = pt.RowRange.Rows.Count
    
    ' 建立樞紐圖表（氣泡圖）
    Set chartObj = wsPivot.ChartObjects.Add( _
        Left:=wsPivot.Range("A1").Left, _
        Top:=wsPivot.Range("A1").Top + pivotRange.Height + 20, _
        Width:=520, _
        Height:=360)
    
    Set cht = chartObj.Chart
    cht.ChartType = xlBubble
    
    ' 使用樞紐結果繪製氣泡圖
    If rowCount > 1 Then
        ' 清除預設數列
        Do While cht.SeriesCollection.Count > 0
            cht.SeriesCollection(1).Delete
        Loop
        
        ' 建立氣泡圖數列
        cht.SeriesCollection.NewSeries
        cht.SeriesCollection(1).Name = "產品分析"
        cht.SeriesCollection(1).XValues = wsPivot.Range("B2:B" & rowCount)
        cht.SeriesCollection(1).Values = wsPivot.Range("C2:C" & rowCount)
        cht.SeriesCollection(1).BubbleSizes = "=" & wsPivot.Name & "!" & "D2:D" & rowCount
    End If
    
    ' 設定圖表標題
    cht.HasTitle = True
    cht.ChartTitle.Text = "樞紐分析氣泡圖 - 產品銷售分析"
    
    ' 設定軸標題
    With cht.Axes(xlCategory)
        .HasTitle = True
        .AxisTitle.Text = "銷售額"
    End With
    
    With cht.Axes(xlValue)
        .HasTitle = True
        .AxisTitle.Text = "利潤"
    End With
    
    ' 設定圖例
    cht.HasLegend = True
    
    ' 套用樣式
    cht.ChartStyle = 8
    
    wsPivot.Activate
    
    MsgBox "樞紐分析氣泡圖已建立完成！" & vbCrLf & _
           "氣泡大小代表市佔率，X軸為銷售額，Y軸為利潤。", vbInformation, "完成"
End Sub

' 填入氣泡圖示範資料
Private Sub FillBubbleChartData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品"
    ws.Range("B1").Value = "銷售額"
    ws.Range("C1").Value = "利潤"
    ws.Range("D1").Value = "市佔率"
    
    ws.Range("A2").Value = "產品A"
    ws.Range("B2").Value = 85000
    ws.Range("C2").Value = 25000
    ws.Range("D2").Value = 0.15
    
    ws.Range("A3").Value = "產品B"
    ws.Range("B3").Value = 62000
    ws.Range("C3").Value = 18000
    ws.Range("D3").Value = 0.08
    
    ws.Range("A4").Value = "產品C"
    ws.Range("B4").Value = 140000
    ws.Range("C4").Value = 45000
    ws.Range("D4").Value = 0.35
    
    ws.Range("A5").Value = "產品D"
    ws.Range("B5").Value = 45000
    ws.Range("C5").Value = 10000
    ws.Range("D5").Value = 0.05
    
    ws.Range("A6").Value = "產品E"
    ws.Range("B6").Value = 110000
    ws.Range("C6").Value = 38000
    ws.Range("D6").Value = 0.25
    
    ws.Range("A7").Value = "產品F"
    ws.Range("B7").Value = 32000
    ws.Range("C7").Value = 8000
    ws.Range("D7").Value = 0.03
    
    ws.Columns("A:D").AutoFit
End Sub
""")

print("Done. 14 sample VBA files generated.")
