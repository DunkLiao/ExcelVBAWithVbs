Attribute VB_Name = "CleanUrlData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanUrlData
'功能說明: 自動清理工作表中 URL 欄位資料，去除多餘空白、修正協定前綴的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************
' 範例使用入口
Sub TestCleanUrlData()
    Call CleanUrlColumn(ActiveSheet, 1)
End Sub

' 清理指定欄位的 URL 資料
' ws       : 要處理的工作表
' colIndex : URL 所在欄索引（從 1 開始）
Sub CleanUrlColumn(ByVal ws As Worksheet, ByVal colIndex As Integer)
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim r As Long
    Dim cellVal As String
    Dim cleaned As String
    Dim lowerCleaned As String
    Dim changedCount As Long

    lastRow = ws.Cells(ws.Rows.Count, colIndex).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "欄位中沒有資料（第 2 列起）。", vbInformation, "無資料"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    changedCount = 0

    For r = 2 To lastRow
        cellVal = CStr(ws.Cells(r, colIndex).Value)

        If Trim(cellVal) = "" Then GoTo NextRow

        cleaned = cellVal

        ' 1. 去除前後空白
        cleaned = Trim(cleaned)

        ' 2. 移除全形空白
        cleaned = Replace(cleaned, Chr(12288), "")

        ' 3. 移除 URL 中的一般空白
        cleaned = Replace(cleaned, " ", "")

        ' 4. 補上 https:// 前綴（若缺少）
        lowerCleaned = LCase(cleaned)
        If Left(lowerCleaned, 7) <> "http://" And Left(lowerCleaned, 8) <> "https://" Then
            If Left(cleaned, 2) = "//" Then
                cleaned = "https:" & cleaned
            Else
                cleaned = "https://" & cleaned
            End If
        End If

        ' 5. 將 http:// 升級為 https://
        If LCase(Left(cleaned, 7)) = "http://" Then
            cleaned = "https://" & Mid(cleaned, 8)
        End If

        ' 6. 移除 URL 尾端的斜線
        Do While Right(cleaned, 1) = "/"
            cleaned = Left(cleaned, Len(cleaned) - 1)
        Loop

        ' 若有修改則寫回並標示
        If cleaned <> cellVal Then
            ws.Cells(r, colIndex).Value = cleaned
            ws.Cells(r, colIndex).Interior.Color = RGB(198, 239, 206)
            changedCount = changedCount + 1
        End If

NextRow:
    Next r

    Application.ScreenUpdating = True
    MsgBox "URL 清理完成！" & vbCrLf & "共修正 " & changedCount & " 筆資料（綠色標示）。", _
           vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清理 URL 時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
