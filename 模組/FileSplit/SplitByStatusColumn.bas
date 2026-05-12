Option Explicit
Attribute VB_Name = "SplitByStatusColumn"
'*************************************************************************************
'模組名稱: SplitByStatusColumn
'功能說明: 依狀態欄位的值，將資料分割為多張獨立工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

Sub SplitByStatusColumn()
    Dim wsSrc As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim statusVal As String
    Dim statusCol As Integer
    Dim headerRow As Long
    Dim dict As Object
    Dim destRow As Long
    Dim safeSheetName As String
    Dim key As Variant

    On Error GoTo ErrHandler

    Set wsSrc = ThisWorkbook.ActiveSheet
    statusCol = 3
    headerRow = 1
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    If lastRow <= headerRow Then
        MsgBox "工作表內沒有資料。", vbExclamation, "提示"
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")

    For i = headerRow + 1 To lastRow
        statusVal = Trim(CStr(wsSrc.Cells(i, statusCol).Value))
        If statusVal = "" Then statusVal = "（空白）"

        If Not dict.Exists(statusVal) Then
            safeSheetName = Left(statusVal, 25)

            On Error Resume Next
            Set wsNew = ThisWorkbook.Sheets(safeSheetName)
            On Error GoTo ErrHandler

            If wsNew Is Nothing Then
                Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                wsNew.Name = safeSheetName
                wsSrc.Rows(headerRow).Copy Destination:=wsNew.Rows(1)
            End If

            dict.Add statusVal, wsNew
            Set wsNew = Nothing
        End If
    Next i

    For i = headerRow + 1 To lastRow
        statusVal = Trim(CStr(wsSrc.Cells(i, statusCol).Value))
        If statusVal = "" Then statusVal = "（空白）"

        Set wsNew = dict(statusVal)
        destRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row + 1
        wsSrc.Rows(i).Copy Destination:=wsNew.Rows(destRow)
    Next i

    For Each key In dict.Keys
        dict(key).Columns.AutoFit
    Next key

    MsgBox "已依狀態欄位完成分割，共建立 " & dict.Count & " 張工作表。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "錯誤：" & Err.Description, vbCritical, "依狀態分割失敗"
End Sub