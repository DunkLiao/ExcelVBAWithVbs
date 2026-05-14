Attribute VB_Name = "CompareByChecksum"
Option Explicit
'*************************************************************************************
'模組名稱: CompareByChecksum
'功能說明: 以列校驗和（字串串接）方式快速比對兩工作表的資料差異
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點
Sub TestCompareByChecksum()
    Call SetupChecksumSampleData
    Call CompareSheetsByChecksum("版本A", "版本B", "校驗差異報告")
End Sub

' 以列校驗和比對兩工作表資料並輸出差異報告
' sheet1Name: 基準工作表名稱
' sheet2Name: 比對工作表名稱
' reportSheetName: 差異報告輸出工作表名稱
Sub CompareSheetsByChecksum(ByVal sheet1Name As String, ByVal sheet2Name As String, _
                             ByVal reportSheetName As String)
    On Error GoTo ErrorHandler

    Dim ws1    As Worksheet
    Dim ws2    As Worksheet
    Dim wsRpt  As Worksheet
    Dim last1  As Long
    Dim last2  As Long
    Dim maxRow As Long
    Dim i      As Long
    Dim ck1    As String
    Dim ck2    As String
    Dim rptRow As Long
    Dim colCnt As Long
    Dim status As String
    Dim desc   As String

    Set ws1 = ThisWorkbook.Worksheets(sheet1Name)
    Set ws2 = ThisWorkbook.Worksheets(sheet2Name)
    Set wsRpt = GetOrCreateChecksumSheet(reportSheetName)
    wsRpt.Cells.Clear

    last1 = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
    last2 = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row
    colCnt = ws1.Cells(1, ws1.Columns.Count).End(xlToLeft).Column
    maxRow = last1
    If last2 > maxRow Then maxRow = last2

    wsRpt.Range("A1:E1").Value = Array("列號", "版本A校驗和", "版本B校驗和", "狀態", "說明")
    With wsRpt.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    rptRow = 2

    Application.ScreenUpdating = False

    For i = 2 To maxRow
        If i <= last1 Then
            ck1 = RowChecksum(ws1, i, colCnt)
        Else
            ck1 = ""
        End If

        If i <= last2 Then
            ck2 = RowChecksum(ws2, i, colCnt)
        Else
            ck2 = ""
        End If

        If ck1 = "" And ck2 <> "" Then
            status = "新增"
            desc = "版本B新增了此列"
        ElseIf ck1 <> "" And ck2 = "" Then
            status = "刪除"
            desc = "版本B刪除了此列"
        ElseIf ck1 <> ck2 Then
            status = "變更"
            desc = "列內容已修改"
        Else
            status = "相同"
            desc = ""
        End If

        If status <> "相同" Then
            wsRpt.Cells(rptRow, 1).Value = i
            wsRpt.Cells(rptRow, 2).Value = ck1
            wsRpt.Cells(rptRow, 3).Value = ck2
            wsRpt.Cells(rptRow, 4).Value = status
            wsRpt.Cells(rptRow, 5).Value = desc

            Select Case status
                Case "新增"
                    wsRpt.Rows(rptRow).Interior.Color = RGB(198, 239, 206)
                Case "刪除"
                    wsRpt.Rows(rptRow).Interior.Color = RGB(255, 199, 206)
                Case "變更"
                    wsRpt.Rows(rptRow).Interior.Color = RGB(255, 235, 156)
            End Select

            rptRow = rptRow + 1
        End If
    Next i

    wsRpt.Columns("A:E").AutoFit
    Application.ScreenUpdating = True

    If rptRow = 2 Then
        MsgBox "兩個工作表資料完全相同！", vbInformation, "比對結果"
    Else
        MsgBox "比對完成！發現 " & rptRow - 2 & " 列差異，詳見「" & reportSheetName & "」。", _
               vbInformation, "比對結果"
    End If
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "校驗和比對時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 計算工作表某列的校驗和字串
Private Function RowChecksum(ByVal ws As Worksheet, ByVal rowIdx As Long, _
                              ByVal colCount As Long) As String
    Dim c      As Long
    Dim result As String
    result = ""
    For c = 1 To colCount
        result = result & CStr(ws.Cells(rowIdx, c).Value) & "|"
    Next c
    RowChecksum = result
End Function

' 建立校驗和比對範例資料
Private Sub SetupChecksumSampleData()
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet

    Set ws1 = GetOrCreateChecksumSheet("版本A")
    ws1.Cells.Clear
    ws1.Range("A1:C1").Value = Array("編號", "品名", "數量")
    ws1.Range("A2:C2").Value = Array(1, "蘋果", 100)
    ws1.Range("A3:C3").Value = Array(2, "香蕉", 200)
    ws1.Range("A4:C4").Value = Array(3, "橘子", 150)
    ws1.Range("A5:C5").Value = Array(4, "葡萄", 80)

    Set ws2 = GetOrCreateChecksumSheet("版本B")
    ws2.Cells.Clear
    ws2.Range("A1:C1").Value = Array("編號", "品名", "數量")
    ws2.Range("A2:C2").Value = Array(1, "蘋果", 120)
    ws2.Range("A3:C3").Value = Array(2, "香蕉", 200)
    ws2.Range("A4:C4").Value = Array(3, "橘子", 150)
    ws2.Range("A5:C5").Value = Array(5, "芒果", 90)
End Sub

' 取得或建立工作表
Private Function GetOrCreateChecksumSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateChecksumSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateChecksumSheet Is Nothing Then
        Set GetOrCreateChecksumSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateChecksumSheet.Name = sheetName
    End If
End Function
