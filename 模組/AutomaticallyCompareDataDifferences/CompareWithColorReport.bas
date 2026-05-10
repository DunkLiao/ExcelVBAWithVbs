Attribute VB_Name = "CompareWithColorReport"
Option Explicit

' ============================================================
' 模組名稱：CompareWithColorReport
' 功能說明：比較兩個工作表資料，輸出色彩標示的差異報告
'           綠色 = 僅在 A 工作表，紅色 = 僅在 B 工作表，
'           黃色 = 兩者皆有但值不同
' 使用方式：活頁簿中需至少有兩個工作表，執行後輸入工作表名稱
' ============================================================

Sub CompareWithColorReport()
    Dim wsA         As Worksheet
    Dim wsB         As Worksheet
    Dim wsReport    As Worksheet
    Dim nameA       As String
    Dim nameB       As String
    Dim reportName  As String
    Dim keyColA     As Long
    Dim lastRowA    As Long
    Dim lastRowB    As Long
    Dim lastColA    As Long
    Dim lastColB    As Long
    Dim i           As Long
    Dim j           As Long
    Dim nextRow     As Long
    
    On Error GoTo ErrHandler
    
    ' 輸入兩個工作表名稱
    nameA = InputBox("請輸入第一個工作表名稱（基準表）：", "工作表A", "Sheet1")
    If nameA = "" Then Exit Sub
    
    nameB = InputBox("請輸入第二個工作表名稱（比較表）：", "工作表B", "Sheet2")
    If nameB = "" Then Exit Sub
    
    ' 驗證工作表存在
    On Error Resume Next
    Set wsA = ThisWorkbook.Sheets(nameA)
    Set wsB = ThisWorkbook.Sheets(nameB)
    On Error GoTo ErrHandler
    
    If wsA Is Nothing Then
        MsgBox "找不到工作表「" & nameA & "」。", vbExclamation, "錯誤"
        Exit Sub
    End If
    If wsB Is Nothing Then
        MsgBox "找不到工作表「" & nameB & "」。", vbExclamation, "錯誤"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' 建立報告工作表
    reportName = "色彩比較報告"
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(reportName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True
    
    Set wsReport = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsReport.Name = reportName
    
    ' 報告標題
    wsReport.Range("A1:E1").Value = Array("差異類型", "列號(A)", "列號(B)", "欄位", "值差異說明")
    With wsReport.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With
    nextRow = 2
    
    lastRowA = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    lastRowB = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
    lastColA = wsA.Cells(1, wsA.Columns.Count).End(xlToLeft).Column
    lastColB = wsB.Cells(1, wsB.Columns.Count).End(xlToLeft).Column
    
    ' 使用第一欄作為鍵值逐列比對
    Dim dictA As Object
    Dim dictB As Object
    Set dictA = CreateObject("Scripting.Dictionary")
    Set dictB = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRowA
        Dim keyA As String
        keyA = CStr(wsA.Cells(i, 1).Value)
        If keyA <> "" Then dictA(keyA) = i
    Next i
    
    For i = 2 To lastRowB
        Dim keyB As String
        keyB = CStr(wsB.Cells(i, 1).Value)
        If keyB <> "" Then dictB(keyB) = i
    Next i
    
    ' 比對：A 有 B 無（綠色）
    Dim k As Variant
    For Each k In dictA.Keys
        If Not dictB.Exists(k) Then
            wsReport.Cells(nextRow, 1).Value = "僅在 " & nameA
            wsReport.Cells(nextRow, 2).Value = dictA(k)
            wsReport.Cells(nextRow, 3).Value = "-"
            wsReport.Cells(nextRow, 4).Value = wsA.Cells(1, 1).Value
            wsReport.Cells(nextRow, 5).Value = "鍵值「" & k & "」"
            wsReport.Rows(nextRow).Interior.Color = RGB(198, 239, 206)
            nextRow = nextRow + 1
        Else
            ' 兩者皆有：逐欄比對差異（黃色）
            Dim rowA As Long, rowB As Long
            rowA = dictA(k)
            rowB = dictB(k)
            Dim maxCol As Long
            maxCol = IIf(lastColA > lastColB, lastColA, lastColB)
            For j = 1 To maxCol
                Dim valA As String, valB As String
                valA = CStr(wsA.Cells(rowA, j).Value)
                valB = CStr(wsB.Cells(rowB, j).Value)
                If valA <> valB Then
                    Dim colName As String
                    colName = CStr(wsA.Cells(1, j).Value)
                    wsReport.Cells(nextRow, 1).Value = "值不同"
                    wsReport.Cells(nextRow, 2).Value = rowA
                    wsReport.Cells(nextRow, 3).Value = rowB
                    wsReport.Cells(nextRow, 4).Value = colName
                    wsReport.Cells(nextRow, 5).Value = nameA & "=" & valA & " / " & nameB & "=" & valB
                    wsReport.Rows(nextRow).Interior.Color = RGB(255, 235, 156)
                    nextRow = nextRow + 1
                End If
            Next j
        End If
    Next k
    
    ' 比對：B 有 A 無（紅色）
    For Each k In dictB.Keys
        If Not dictA.Exists(k) Then
            wsReport.Cells(nextRow, 1).Value = "僅在 " & nameB
            wsReport.Cells(nextRow, 2).Value = "-"
            wsReport.Cells(nextRow, 3).Value = dictB(k)
            wsReport.Cells(nextRow, 4).Value = wsB.Cells(1, 1).Value
            wsReport.Cells(nextRow, 5).Value = "鍵值「" & k & "」"
            wsReport.Rows(nextRow).Interior.Color = RGB(255, 199, 206)
            nextRow = nextRow + 1
        End If
    Next k
    
    wsReport.Columns.AutoFit
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    Dim diffCount As Long
    diffCount = nextRow - 2
    MsgBox "色彩比較報告建立完成！" & vbCrLf & _
           "共發現 " & diffCount & " 筆差異。" & vbCrLf & _
           "請查看「" & reportName & "」工作表。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub