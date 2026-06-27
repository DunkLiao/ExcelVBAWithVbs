Attribute VB_Name = "MergeExcelWithDedupSummary"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelWithDedupSummary
'功能說明: 合併多個Excel檔案並自動去重，同時產生合併摘要報告的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestMergeExcelWithDedupSummary()
    Call MergeExcelWithDedupSummary("C:\Temp\MergeData")
End Sub

Sub MergeExcelWithDedupSummary(ByVal folderPath As String)
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim summaryWs As Worksheet
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim lastRow As Long
    Dim sourceLastRow As Long
    Dim j As Long
    Dim totalAdded As Long
    Dim totalSkipped As Long
    Dim fileName As String
    Dim keyVal As String
    Dim existingKeys As Object
    Dim fso As Object
    Dim folderObj As Object
    Dim fileObj As Object
    Dim fileCount As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        MsgBox "資料夾不存在：" & folderPath, vbExclamation, "錯誤"
        Set fso = Nothing
        Exit Sub
    End If
    
    Set targetWb = Workbooks.Add
    Set targetWs = targetWb.Worksheets(1)
    targetWs.Name = "合併結果"
    
    targetWs.Range("A1").Value = "編號"
    targetWs.Range("B1").Value = "名稱"
    targetWs.Range("C1").Value = "金額"
    targetWs.Range("D1").Value = "來源"
    
    lastRow = 1
    totalAdded = 0
    totalSkipped = 0
    fileCount = 0
    Set existingKeys = CreateObject("Scripting.Dictionary")
    
    Set folderObj = fso.GetFolder(folderPath)
    
    For Each fileObj In folderObj.Files
        If LCase(fso.GetExtensionName(fileObj.Name)) = "xlsx" Then
            fileCount = fileCount + 1
            fileName = fileObj.Name
            
            Set sourceWb = Workbooks.Open(fileObj.Path)
            Set sourceWs = sourceWb.Worksheets(1)
            sourceLastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row
            
            If sourceLastRow > 1 Then
                For j = 2 To sourceLastRow
                    keyVal = CStr(sourceWs.Cells(j, 1).Value)
                    
                    If Not existingKeys.Exists(keyVal) Then
                        existingKeys.Add keyVal, True
                        lastRow = lastRow + 1
                        targetWs.Cells(lastRow, 1).Value = sourceWs.Cells(j, 1).Value
                        targetWs.Cells(lastRow, 2).Value = sourceWs.Cells(j, 2).Value
                        targetWs.Cells(lastRow, 3).Value = sourceWs.Cells(j, 3).Value
                        targetWs.Cells(lastRow, 4).Value = fileName
                        totalAdded = totalAdded + 1
                    Else
                        totalSkipped = totalSkipped + 1
                    End If
                Next j
            End If
            
            sourceWb.Close False
        End If
    Next fileObj
    
    Set summaryWs = targetWb.Worksheets.Add
    summaryWs.Name = "合併摘要"
    summaryWs.Range("A1").Value = "合併摘要報告"
    summaryWs.Range("A1").Font.Bold = True
    summaryWs.Range("A3").Value = "處理檔案數"
    summaryWs.Range("B3").Value = fileCount
    summaryWs.Range("A4").Value = "新增筆數"
    summaryWs.Range("B4").Value = totalAdded
    summaryWs.Range("A5").Value = "略過重複"
    summaryWs.Range("B5").Value = totalSkipped
    summaryWs.Range("A6").Value = "合計總筆數"
    summaryWs.Range("B6").Value = totalAdded + totalSkipped
    
    summaryWs.Columns("A:B").AutoFit
    targetWs.Columns("A:D").AutoFit
    targetWs.Activate
    
    Set existingKeys = Nothing
    Set fso = Nothing
    
    MsgBox "檔案合併去重完成！" & vbCrLf & _
           "新增: " & totalAdded & " 筆" & vbCrLf & _
           "略過重複: " & totalSkipped & " 筆", vbInformation, "完成"
End Sub
