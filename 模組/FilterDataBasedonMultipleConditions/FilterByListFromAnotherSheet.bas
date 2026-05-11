Attribute VB_Name = "FilterByListFromAnotherSheet"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByListFromAnotherSheet
'功能說明: 讀取另一工作表的清單作為篩選條件，將資料表中符合清單的列複製至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestFilterByListFromAnotherSheet()
    Call FilterByListFromSheet
End Sub

' 依另一工作表的清單篩選資料
Sub FilterByListFromSheet()
    On Error GoTo ErrorHandler

    Dim wsData As Worksheet
    Dim wsList As Worksheet
    Dim wsResult As Worksheet
    Dim dataSheetName As String
    Dim listSheetName As String
    Dim dataColIdx As Integer
    Dim lastDataRow As Long
    Dim lastListRow As Long
    Dim r As Long
    Dim lr As Long
    Dim targetRow As Long
    Dim cellVal As String
    Dim listVal As String
    Dim lastDataCol As Long
    Dim filterDict As Object
    Dim dataColStr As String

    ' 設定工作表名稱
    dataSheetName = InputBox("請輸入資料工作表的名稱：", "設定資料工作表", "資料")
    If dataSheetName = "" Then
        MsgBox "已取消操作", vbInformation, "取消"
        Exit Sub
    End If

    listSheetName = InputBox("請輸入篩選清單工作表的名稱：", "設定清單工作表", "篩選清單")
    If listSheetName = "" Then
        MsgBox "已取消操作", vbInformation, "取消"
        Exit Sub
    End If

    ' 取得工作表
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets(dataSheetName)
    Set wsList = ThisWorkbook.Worksheets(listSheetName)
    On Error GoTo ErrorHandler

    If wsData Is Nothing Then
        MsgBox "找不到工作表：" & dataSheetName, vbExclamation, "錯誤"
        Exit Sub
    End If
    If wsList Is Nothing Then
        MsgBox "找不到工作表：" & listSheetName, vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 設定比對欄位索引
    dataColStr = InputBox("請輸入資料工作表中要比對的欄號（數字，例如 1 表示 A 欄）：", _
                          "設定比對欄位", "1")
    If dataColStr = "" Then
        MsgBox "已取消操作", vbInformation, "取消"
        Exit Sub
    End If
    dataColIdx = CInt(dataColStr)

    ' 建立篩選清單字典
    Set filterDict = CreateObject("Scripting.Dictionary")
    filterDict.CompareMode = 1 ' vbTextCompare

    lastListRow = wsList.Cells(wsList.Rows.Count, 1).End(xlUp).Row
    For lr = 1 To lastListRow
        listVal = Trim(CStr(wsList.Cells(lr, 1).Value))
        If listVal <> "" And Not filterDict.Exists(listVal) Then
            filterDict.Add listVal, True
        End If
    Next lr

    If filterDict.Count = 0 Then
        MsgBox "篩選清單為空！", vbExclamation, "無清單資料"
        Exit Sub
    End If

    ' 建立或清空結果工作表
    On Error Resume Next
    Set wsResult = ThisWorkbook.Worksheets("清單篩選結果")
    On Error GoTo ErrorHandler
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Worksheets.Add
        wsResult.Name = "清單篩選結果"
    Else
        wsResult.Cells.Clear
    End If

    Application.ScreenUpdating = False

    lastDataRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastDataCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column

    ' 複製標題列
    wsData.Range(wsData.Cells(1, 1), wsData.Cells(1, lastDataCol)).Copy _
        Destination:=wsResult.Cells(1, 1)
    wsResult.Rows(1).Font.Bold = True
    targetRow = 2

    ' 逐列比對
    For r = 2 To lastDataRow
        cellVal = Trim(CStr(wsData.Cells(r, dataColIdx).Value))
        If filterDict.Exists(cellVal) Then
            wsData.Range(wsData.Cells(r, 1), wsData.Cells(r, lastDataCol)).Copy _
                Destination:=wsResult.Cells(targetRow, 1)
            targetRow = targetRow + 1
        End If
    Next r

    wsResult.UsedRange.Columns.AutoFit
    Application.ScreenUpdating = True
    wsResult.Activate

    MsgBox "篩選完成！" & vbCrLf & _
           "清單項目數：" & filterDict.Count & " 個" & vbCrLf & _
           "符合的資料：" & targetRow - 2 & " 筆", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "篩選時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
