Attribute VB_Name = "MergeExcelWithFooterRow"
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
