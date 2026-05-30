Attribute VB_Name = "SplitByConditionalGroup"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByConditionalGroup
'功能說明: 依條件群組（數值區間分組）將工作表資料切割為多個子工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/30
'
'*************************************************************************************

' 程式進入點
Sub TestSplitByConditionalGroup()
    Call SplitDataByConditionalGroup
End Sub

' 依條件群組切割資料
Sub SplitDataByConditionalGroup()
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim lngLastRow As Long
    Dim lngDestRow As Long
    Dim i As Long
    Dim lngScore As Long
    Dim sGroup As String

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False

    Set wsSrc = ActiveSheet
    lngLastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row

    If lngLastRow < 2 Then
        MsgBox "資料不足，至少需要一列資料（不含標題）。", vbExclamation
        GoTo CleanUp
    End If

    Call DeleteOldGroupSheets(ThisWorkbook)

    For i = 2 To lngLastRow
        lngScore = 0
        On Error Resume Next
        lngScore = CLng(wsSrc.Cells(i, 2).Value)
        On Error GoTo ErrHandler
        sGroup = GetScoreGroupName(lngScore)
        Set wsDest = GetOrCreateGroupSheet(ThisWorkbook, wsSrc, sGroup)
        lngDestRow = wsDest.Cells(wsDest.Rows.Count, 1).End(xlUp).Row + 1
        wsSrc.Rows(i).Copy Destination:=wsDest.Rows(lngDestRow)
    Next i

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> wsSrc.Name Then
            ws.Columns.AutoFit
        End If
    Next ws

    Application.ScreenUpdating = True
    MsgBox "依條件群組切割完成！", vbInformation, "完成"
    Exit Sub

CleanUp:
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 依數值決定群組名稱
Private Function GetScoreGroupName(ByVal score As Long) As String
    If score >= 90 Then
        GetScoreGroupName = "群組A_90以上"
    ElseIf score >= 70 Then
        GetScoreGroupName = "群組B_70到89"
    ElseIf score >= 50 Then
        GetScoreGroupName = "群組C_50到69"
    Else
        GetScoreGroupName = "群組D_50以下"
    End If
End Function

' 取得或建立分組工作表，若為新工作表則複製標題列
Private Function GetOrCreateGroupSheet(ByVal wb As Workbook, _
    ByVal wsSrc As Worksheet, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
        wsSrc.Rows(1).Copy Destination:=ws.Rows(1)
        ws.Rows(1).Font.Bold = True
    End If
    Set GetOrCreateGroupSheet = ws
End Function

' 刪除舊的分組工作表
Private Sub DeleteOldGroupSheets(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim wsNames() As String
    Dim nCount As Integer
    Dim i As Integer

    nCount = 0
    For Each ws In wb.Worksheets
        If Left(ws.Name, 2) = "群組" Then
            nCount = nCount + 1
        End If
    Next ws

    If nCount = 0 Then Exit Sub

    ReDim wsNames(0 To nCount - 1)
    i = 0
    For Each ws In wb.Worksheets
        If Left(ws.Name, 2) = "群組" Then
            wsNames(i) = ws.Name
            i = i + 1
        End If
    Next ws

    Application.DisplayAlerts = False
    For i = 0 To nCount - 1
        wb.Worksheets(wsNames(i)).Delete
    Next i
    Application.DisplayAlerts = True
End Sub
