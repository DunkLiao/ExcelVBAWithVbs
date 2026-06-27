Attribute VB_Name = "MergeWithDataTransform"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithDataTransform
'功能說明: 合併多工作表資料時同時進行資料轉換（數值格式化、文字修剪）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestMergeWithDataTransform()
    Call MergeSheetsWithTransform
End Sub

Sub MergeSheetsWithTransform()
    Dim wsResult As Worksheet
    Dim ws As Worksheet
    Dim i As Integer
    Dim lastRow As Long
    Dim destRow As Long

    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsResult = ThisWorkbook.Worksheets("轉換合併結果")
    If Not wsResult Is Nothing Then wsResult.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' 建立測試用工作表
    For i = 1 To 3
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "分店" & i

        ws.Range("A1").Value = "員工"
        ws.Range("B1").Value = "職位"
        ws.Range("C1").Value = "薪資(文字格式)"
        ws.Range("D1").Value = "入職日期"
        ws.Range("A1:D1").Font.Bold = True

        ws.Range("A2").Value = " 員工A  "
        ws.Range("B2").Value = "經理"
        ws.Range("C2").Value = "  $45,000  "
        ws.Range("D2").Value = "2024-03-15"

        ws.Range("A3").Value = " 員工B  "
        ws.Range("B3").Value = "專員"
        ws.Range("C3").Value = "  $32,000  "
        ws.Range("D3").Value = "2024-06-01"

        ws.Columns("A:D").AutoFit
    Next i

    ' 建立合併結果工作表
    Set wsResult = ThisWorkbook.Worksheets.Add
    wsResult.Name = "轉換合併結果"

    ' 標題列
    wsResult.Range("A1").Value = "來源"
    wsResult.Range("B1").Value = "員工"
    wsResult.Range("C1").Value = "職位"
    wsResult.Range("D1").Value = "薪資(數值)"
    wsResult.Range("E1").Value = "入職日期"
    wsResult.Range("A1:E1").Font.Bold = True

    destRow = 2

    ' 遍歷每個工作表
    Dim wsItem As Worksheet
    For Each wsItem In ThisWorkbook.Worksheets
        If Left(wsItem.Name, 2) = "分店" Then
            lastRow = wsItem.Cells(wsItem.Rows.Count, 1).End(xlUp).Row

            Dim r As Long
            For r = 2 To lastRow
                ' 來源工作表標記
                wsResult.Cells(destRow, 1).Value = wsItem.Name

                ' 文字修剪
                wsResult.Cells(destRow, 2).Value = Trim(wsItem.Cells(r, 1).Value)

                ' 職位直接複製
                wsResult.Cells(destRow, 3).Value = Trim(wsItem.Cells(r, 2).Value)

                ' 薪資轉換：去除 $ 和逗號並轉換為數值
                Dim rawStr As String
                rawStr = wsItem.Cells(r, 3).Value
                rawStr = Replace(rawStr, "$", "")
                rawStr = Replace(rawStr, ",", "")
                rawStr = Trim(rawStr)
                If IsNumeric(rawStr) Then
                    wsResult.Cells(destRow, 4).Value = CDbl(rawStr)
                    wsResult.Cells(destRow, 4).NumberFormat = "#,##0"
                End If

                ' 日期直接複製
                wsResult.Cells(destRow, 5).Value = wsItem.Cells(r, 4).Value

                destRow = destRow + 1
            Next r
        End If
    Next wsItem

    wsResult.Columns("A:E").AutoFit

    MsgBox "資料合併與轉換完成！共處理 " & (destRow - 2) & " 筆資料。", vbInformation, "完成"
End Sub
