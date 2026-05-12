Option Explicit
Attribute VB_Name = "FilterAndCopyToNewSheet"
'*************************************************************************************
'模組名稱: FilterAndCopyToNewSheet
'功能說明: 依指定欄位條件篩選資料，並將符合條件的列複製至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub FilterAndCopyToNewSheet()
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim destRow As Long
    Dim i As Long
    Dim filterCol As Integer
    Dim filterValue As String
    Dim copyCount As Long
    Dim destSheetName As String
    Dim cellVal As String

    On Error GoTo ErrHandler

    Set wsSrc = ThisWorkbook.ActiveSheet
    filterCol = 3
    filterValue = "達標"
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "工作表內沒有足夠的資料。", vbExclamation, "提示"
        Exit Sub
    End If

    destSheetName = "篩選結果_" & filterValue

    On Error Resume Next
    Set wsDest = ThisWorkbook.Sheets(destSheetName)
    On Error GoTo ErrHandler

    If wsDest Is Nothing Then
        Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDest.Name = destSheetName
    Else
        wsDest.Cells.Clear
    End If

    wsSrc.Rows(1).Copy Destination:=wsDest.Rows(1)
    destRow = 2
    copyCount = 0

    For i = 2 To lastRow
        cellVal = Trim(CStr(wsSrc.Cells(i, filterCol).Value))
        If cellVal = filterValue Then
            wsSrc.Rows(i).Copy Destination:=wsDest.Rows(destRow)
            destRow = destRow + 1
            copyCount = copyCount + 1
        End If
    Next i

    wsDest.Columns.AutoFit
    wsDest.Activate

    MsgBox "篩選完成！" & vbCrLf & _
           "條件：第 " & filterCol & " 欄 = """ & filterValue & """" & vbCrLf & _
           "共複製 " & copyCount & " 筆資料至「" & destSheetName & "」。", _
           vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "篩選並複製資料失敗"
End Sub