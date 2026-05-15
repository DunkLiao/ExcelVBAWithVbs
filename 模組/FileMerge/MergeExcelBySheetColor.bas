Attribute VB_Name = "MergeExcelBySheetColor"
Option Explicit
'*************************************************************************************
'模組名稱: MergeExcelBySheetColor
'功能說明: 依索引標籤顏色合併符合條件的工作表資料到合併結果工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

Public Sub RunMergeExcelBySheetColor()
    On Error GoTo ErrorHandler

    Dim targetColor As Long
    Dim colorLabel As String
    Dim wsResult As Worksheet
    Dim ws As Worksheet
    Dim destRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim isFirstSheet As Boolean
    Dim mergedSheetCount As Long

    If Not GetSelectedTabColor(targetColor, colorLabel) Then Exit Sub

    Set wsResult = GetOrCreateColorMergeSheet("合併結果")
    wsResult.Cells.Clear

    destRow = 1
    isFirstSheet = True
    mergedSheetCount = 0

    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsResult.Name Then
            If SheetMatchesTabColor(ws, targetColor) Then
                lastRow = GetLastUsedRow(ws)
                lastCol = GetLastUsedCol(ws)

                If lastRow >= 1 And lastCol >= 1 Then
                    If isFirstSheet Then
                        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Copy _
                            Destination:=wsResult.Cells(destRow, 1)
                        destRow = destRow + lastRow
                        isFirstSheet = False
                    ElseIf lastRow > 1 Then
                        ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol)).Copy _
                            Destination:=wsResult.Cells(destRow, 1)
                        destRow = destRow + lastRow - 1
                    End If
                    mergedSheetCount = mergedSheetCount + 1
                End If
            End If
        End If
    Next ws

    Application.ScreenUpdating = True
    wsResult.Columns.AutoFit

    If mergedSheetCount = 0 Then
        MsgBox "找不到索引標籤顏色為 " & colorLabel & " 的工作表。", vbInformation, "提示"
    Else
        MsgBox "已合併 " & mergedSheetCount & " 張 " & colorLabel & " 工作表。", vbInformation, "完成"
    End If
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "依顏色合併工作表時發生錯誤: " & Err.Description, vbExclamation, "錯誤"
End Sub

Private Function GetSelectedTabColor(ByRef targetColor As Long, ByRef colorLabel As String) As Boolean
    Dim userChoice As String

    userChoice = Trim$(InputBox( _
        Prompt:="請選擇索引標籤顏色:" & vbCrLf & _
                "1 = 紅色" & vbCrLf & _
                "2 = 藍色" & vbCrLf & _
                "3 = 綠色" & vbCrLf & _
                "4 = 黃色", _
        Title:="選擇顏色", _
        Default:="1"))

    If userChoice = "" Then Exit Function

    Select Case userChoice
        Case "1"
            targetColor = RGB(255, 0, 0)
            colorLabel = "紅色"
        Case "2"
            targetColor = RGB(0, 112, 192)
            colorLabel = "藍色"
        Case "3"
            targetColor = RGB(0, 176, 80)
            colorLabel = "綠色"
        Case "4"
            targetColor = RGB(255, 255, 0)
            colorLabel = "黃色"
        Case Else
            MsgBox "請輸入 1 到 4 的數字。", vbExclamation, "提示"
            Exit Function
    End Select

    GetSelectedTabColor = True
End Function

Private Function SheetMatchesTabColor(ByVal ws As Worksheet, ByVal targetColor As Long) As Boolean
    On Error Resume Next
    If ws.Tab.ColorIndex <> xlColorIndexNone Then
        SheetMatchesTabColor = (CLng(ws.Tab.Color) = targetColor)
    End If
    On Error GoTo 0
End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        GetLastUsedRow = 0
    Else
        GetLastUsedRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
End Function

Private Function GetLastUsedCol(ByVal ws As Worksheet) As Long
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        GetLastUsedCol = 0
    Else
        GetLastUsedCol = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
End Function

Private Function GetOrCreateColorMergeSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateColorMergeSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateColorMergeSheet Is Nothing Then
        Set GetOrCreateColorMergeSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateColorMergeSheet.Name = sheetName
    End If
End Function
