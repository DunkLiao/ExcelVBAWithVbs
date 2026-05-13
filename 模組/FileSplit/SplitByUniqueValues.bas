Attribute VB_Name = "SplitByUniqueValues"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByUniqueValues
'功能說明: 依指定欄位的唯一值，將工作表資料分割至多個新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Sub SplitByUniqueValues()
    Dim srcWs       As Worksheet
    Dim dstWs       As Worksheet
    Dim colIdx      As Long
    Dim lastRow     As Long
    Dim i           As Long
    Dim keyVal      As String
    Dim dict        As Object
    Dim key         As Variant

    Set srcWs = ActiveSheet
    lastRow = srcWs.Cells(srcWs.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "資料不足，無法分割。", vbExclamation, "提示"
        Exit Sub
    End If

    Dim colInput As String
    colInput = InputBox("請輸入要依據分割的欄號（例如：1 代表 A 欄）：", "設定分割欄位", "1")
    If colInput = "" Then Exit Sub
    colIdx = CLng(colInput)

    Set dict = CreateObject("Scripting.Dictionary")
    For i = 2 To lastRow
        keyVal = CStr(srcWs.Cells(i, colIdx).Value)
        If keyVal <> "" Then
            If Not dict.exists(keyVal) Then
                dict.Add keyVal, Nothing
            End If
        End If
    Next i

    If dict.Count = 0 Then
        MsgBox "指定欄位沒有找到任何唯一值。", vbExclamation, "提示"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For Each key In dict.Keys
        Set dstWs = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

        Dim wsName As String
        wsName = Left(CStr(key), 31)
        On Error Resume Next
        dstWs.Name = wsName
        On Error GoTo 0

        srcWs.Rows(1).Copy Destination:=dstWs.Rows(1)

        Dim destRow As Long
        destRow = 2
        For i = 2 To lastRow
            If CStr(srcWs.Cells(i, colIdx).Value) = CStr(key) Then
                srcWs.Rows(i).Copy Destination:=dstWs.Rows(destRow)
                destRow = destRow + 1
            End If
        Next i

        dstWs.Columns.AutoFit
    Next key

    Application.ScreenUpdating = True
    MsgBox "分割完成，共建立 " & dict.Count & " 個工作表。", vbInformation, "完成"
End Sub
