Attribute VB_Name = "ClearNonStandardFormatting"
Option Explicit
'*************************************************************************************
'模組名稱: ClearNonStandardFormatting
'功能說明: 清除非標準格式的儲存格，僅保留指定的字型、大小、顏色、框線等標準格式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestClearNonStandardFormatting()
    Call ClearNonStandardFormatting
End Sub

' 清除非標準格式，只保留指定格式
Sub ClearNonStandardFormatting()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim targetRange As Range
    Dim cell As Range
    Dim lastRow As Long
    Dim lastCol As Long
    
    sheetName = "清除非標準格式"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillClearFormatData(ws)
    
    ' 顯示清除前的狀態
    MsgBox "即將清除非標準格式。" & vbCrLf & _
           "標準格式定義：字型=微軟正黑體, 大小=11, 字型色=黑色", vbInformation, "提示"
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Set targetRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    
    For Each cell In targetRange
        With cell
            ' 清除非標準字型
            If .Font.Name <> "微軟正黑體" Then
                .Font.Name = "微軟正黑體"
            End If
            
            ' 清除非標準字型大小
            If .Font.Size <> 11 Then
                .Font.Size = 11
            End If
            
            ' 清除非標準字型色（非黑色）
            If .Font.Color <> RGB(0, 0, 0) Then
                .Font.Color = RGB(0, 0, 0)
            End If
            
            ' 清除非標準粗體
            If .Font.Bold = True Then
                .Font.Bold = False
            End If
            
            ' 清除非標準斜體
            If .Font.Italic = True Then
                .Font.Italic = False
            End If
            
            ' 清除非標準底線
            If .Font.Underline <> xlUnderlineStyleNone Then
                .Font.Underline = xlUnderlineStyleNone
            End If
            
            ' 清除非白色/無填滿的背景色
            If .Interior.Color <> RGB(255, 255, 255) And _
               .Interior.ColorIndex <> xlNone Then
                .Interior.ColorIndex = xlNone
            End If
            
            ' 清除內部框線（只保留最外框）
            If .Borders(xlInsideVertical).LineStyle <> xlNone Then
                .Borders(xlInsideVertical).LineStyle = xlNone
            End If
            If .Borders(xlInsideHorizontal).LineStyle <> xlNone Then
                .Borders(xlInsideHorizontal).LineStyle = xlNone
            End If
        End With
    Next cell
    
    ws.Columns("A:C").AutoFit
    ws.Activate
    
    MsgBox "非標準格式已清除完成！", vbInformation, "完成"
End Sub

' 填入示範資料（含多種不同格式）
Private Sub FillClearFormatData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品名稱"
    ws.Range("B1").Value = "數量"
    ws.Range("C1").Value = "備註"
    
    ' 標題列使用特殊格式
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(0, 0, 255)
        .Interior.Color = RGB(200, 200, 255)
    End With
    
    ws.Range("A2").Value = "產品X"
    ws.Range("B2").Value = 100
    ws.Range("C2").Value = "正常"
    ws.Range("A2").Font.Italic = True
    
    ws.Range("A3").Value = "產品Y"
    ws.Range("B3").Value = 200
    ws.Range("C3").Value = "測試"
    ws.Range("A3").Font.Underline = xlUnderlineStyleSingle
    
    ws.Range("A4").Value = "產品Z"
    ws.Range("B4").Value = 300
    ws.Range("C4").Value = "特例"
    ws.Range("A4").Interior.Color = RGB(255, 255, 0)
    
    ws.Range("A5").Value = "產品W"
    ws.Range("B5").Value = 400
    ws.Range("C5").Value = "緊急"
    ws.Range("A5").Font.Color = RGB(255, 0, 0)
    ws.Range("A5").Font.Size = 16
    
    ws.Columns("A:C").AutoFit
End Sub
