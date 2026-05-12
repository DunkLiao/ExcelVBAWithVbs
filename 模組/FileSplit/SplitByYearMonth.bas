Attribute VB_Name = "SplitByYearMonth"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByYearMonth
'功能說明: 依據日期欄位的年月，將工作表資料分割為各別工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/12
'
'*************************************************************************************

' 範例進入點
Sub TestSplitByYearMonth()
    Call SplitByYearMonth
End Sub

' 依年月欄位分割資料
Sub SplitByYearMonth()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dateCol As Long
    Dim colInput As String
    Dim ym As String
    Dim sheetDict As Object
    Dim newWs As Worksheet
    Dim destRow As Long
    Dim i As Long
    Dim cellVal As Variant

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    colInput = InputBox("請輸入日期欄位的欄號 (例如 A、B、C)：", "指定日期欄", "A")
    If colInput = "" Then
        MsgBox "未輸入欄號，程式結束。", vbInformation, "取消"
        Exit Sub
    End If
    dateCol = ws.Range(colInput & "1").Column

    Set sheetDict = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        cellVal = ws.Cells(i, dateCol).Value
        If IsDate(cellVal) Then
            ym = Format(CDate(cellVal), "YYYY-MM")
        ElseIf VarType(cellVal) = vbDouble Then
            ym = Format(CDate(cellVal), "YYYY-MM")
        Else
            ym = "其他"
        End If

        If Not sheetDict.Exists(ym) Then
            On Error Resume Next
            Set newWs = ThisWorkbook.Worksheets(ym)
            On Error GoTo ErrorHandler
            If newWs Is Nothing Then
                Set newWs = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
                newWs.Name = ym
            End If
            newWs.Cells.Clear
            ws.Rows(1).Copy Destination:=newWs.Rows(1)
            newWs.Rows(1).Font.Bold = True
            sheetDict(ym) = 2
            Set newWs = Nothing
        End If

        destRow = sheetDict(ym)
        ws.Rows(i).Copy Destination:=ThisWorkbook.Worksheets(ym).Rows(destRow)
        sheetDict(ym) = destRow + 1
    Next i

    Dim key As Variant
    For Each key In sheetDict.Keys
        ThisWorkbook.Worksheets(CStr(key)).Columns.AutoFit
    Next key

    Application.ScreenUpdating = True
    MsgBox "依年月分割完成，共建立 " & sheetDict.Count & " 個工作表！", vbInformation, "完成"
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "分割時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub
