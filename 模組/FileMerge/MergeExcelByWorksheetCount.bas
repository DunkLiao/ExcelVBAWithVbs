Attribute VB_Name = "MergeExcelByWorksheetCount"
Option Explicit

'*************************************************************************************
'模組名稱: MergeExcelByWorksheetCount
'功能說明: 依工作表數量排序後合並多個 Excel 檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub MergeExcelByWorksheetCount()
    Dim fso As Object
    Dim folderPath As String
    Dim fileArr() As String
    Dim countArr() As Long
    Dim wbSrc As Workbook
    Dim wsDest As Worksheet
    Dim f As Object
    Dim i As Long, j As Long
    Dim tempName As String
    Dim tempCount As Long
    Dim fileCount As Long
    Dim ws As Worksheet
    Dim lastR As Long
    Dim destRow As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    folderPath = ThisWorkbook.Path

    fileCount = 0
    For Each f In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(f.Name)) = "xlsx" Then
            If f.Name <> ThisWorkbook.Name Then
                fileCount = fileCount + 1
                ReDim Preserve fileArr(1 To fileCount)
                ReDim Preserve countArr(1 To fileCount)
                fileArr(fileCount) = f.Path
            End If
        End If
    Next f

    If fileCount = 0 Then
        MsgBox "找不到可合並的 .xlsx 檔案。", vbExclamation
        Exit Sub
    End If

    For i = 1 To fileCount
        Set wbSrc = Workbooks.Open(fileArr(i), ReadOnly:=True)
        countArr(i) = wbSrc.Worksheets.Count
        wbSrc.Close False
    Next i

    For i = 1 To fileCount - 1
        For j = i + 1 To fileCount
            If countArr(j) > countArr(i) Then
                tempCount = countArr(i)
                countArr(i) = countArr(j)
                countArr(j) = tempCount
                tempName = fileArr(i)
                fileArr(i) = fileArr(j)
                fileArr(j) = tempName
            End If
        Next j
    Next i

    Set wsDest = ThisWorkbook.Worksheets.Add
    wsDest.Name = "合並結果"

    destRow = 1
    For i = 1 To fileCount
        Set wbSrc = Workbooks.Open(fileArr(i), ReadOnly:=True)
        For Each ws In wbSrc.Worksheets
            lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            If lastR >= 1 Then
                ws.Range("A1:Z" & lastR).Copy _
                    wsDest.Cells(destRow, 1)
                destRow = destRow + lastR + 1
            End If
        Next ws
        wbSrc.Close False
    Next i

    MsgBox "合並完成，共處理 " & fileCount & " 個檔案！", vbInformation
End Sub
