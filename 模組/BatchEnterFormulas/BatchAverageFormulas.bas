Attribute VB_Name = "BatchAverageFormulas"
Option Explicit

' ============================================================
' 範例：批次在每列的最後一欄插入 AVERAGE 平均值公式
' 功能：自動偵測資料範圍，在每列末尾插入平均公式
' ============================================================
Sub BatchInsertAverageFormulas()
    Dim ws          As Worksheet
    Dim lngLastRow  As Long
    Dim lngLastCol  As Long
    Dim lngRow      As Long
    Dim strColLetter As String

    On Error GoTo ErrHandler
    Set ws = ActiveSheet

    lngLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lngLastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lngLastRow < 2 Or lngLastCol < 2 Then
        MsgBox "資料不足，請確認工作表至少有兩列兩欄資料。", vbExclamation
        Exit Sub
    End If

    ' --- 在每列資料末欄+1插入 AVERAGE 公式 ---
    strColLetter = Split(ws.Cells(1, lngLastCol + 1).Address, "$")(1)

    ' 第一列寫標題
    ws.Cells(1, lngLastCol + 1).Value = "平均值"

    Application.ScreenUpdating = False
    For lngRow = 2 To lngLastRow
        Dim strStartCell As String
        Dim strEndCell   As String
        strStartCell = ws.Cells(lngRow, 2).Address(False, False)
        strEndCell = ws.Cells(lngRow, lngLastCol).Address(False, False)
        ws.Cells(lngRow, lngLastCol + 1).Formula = _
            "=AVERAGE(" & strStartCell & ":" & strEndCell & ")"
    Next lngRow
    Application.ScreenUpdating = True

    MsgBox "已在第 " & lngLastCol + 1 & " 欄批次插入 AVERAGE 公式，共 " & _
        lngLastRow - 1 & " 列。", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "發生錯誤：" & Err.Description, vbCritical
End Sub
