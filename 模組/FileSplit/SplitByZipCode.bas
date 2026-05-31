Attribute VB_Name = "SplitByZipCode"
Option Explicit

'*************************************************************************************
'模組名稱: SplitByZipCode
'功能說明: 依郵遞區號欄位將工作表拆分為多個工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撒寫日期: 2025/6/1
'
'*************************************************************************************

Sub SplitByZipCode()
    Dim wsSrc As Worksheet
    Dim wsNew As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim zipCol As Long
    Dim zipCode As String
    Dim dictKeys As Object
    Dim key As Variant
    Dim destRow As Long

    Set wsSrc = ThisWorkbook.Worksheets(1)
    Set dictKeys = CreateObject("Scripting.Dictionary")

    zipCol = 3
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, zipCol).End(xlUp).Row

    For i = 2 To lastRow
        zipCode = Trim(CStr(wsSrc.Cells(i, zipCol).Value))
        If zipCode <> "" Then
            If Not dictKeys.Exists(zipCode) Then
                dictKeys.Add zipCode, zipCode
            End If
        End If
    Next i

    For Each key In dictKeys.Keys
        On Error Resume Next
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets(CStr(key)).Delete
        Application.DisplayAlerts = True
        On Error GoTo 0

        Set wsNew = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsNew.Name = CStr(key)

        wsSrc.Rows(1).Copy wsNew.Rows(1)

        destRow = 2
        For i = 2 To lastRow
            If Trim(CStr(wsSrc.Cells(i, zipCol).Value)) = CStr(key) Then
                wsSrc.Rows(i).Copy wsNew.Rows(destRow)
                destRow = destRow + 1
            End If
        Next i
    Next key

    MsgBox "拆分完成，共建立 " & dictKeys.Count & " 個工作表！", vbInformation
End Sub
