Option Explicit
'*************************************************************************************
'模組名稱: FilterByTopPercent
'功能說明: 依指定欄位的數值，篩選前 N% 的資料列並複製到新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/27
'
'*************************************************************************************
Sub FilterByTopPercent()
    Dim ws As Worksheet
    Dim newWs As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim colNum As Long
    Dim pct As Double
    Dim threshold As Double
    Dim userInput As String
    Dim colInput As String
    Dim i As Long
    Dim destRow As Long
    Dim pctInt As Long
    Dim newSheetName As String

    On Error GoTo ErrHandler

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        MsgBox "資料不足！需要至少一列標題加一列資料。", vbExclamation, "提示"
        Exit Sub
    End If

    ' 詢問篩選欄位（欄號）
    colInput = InputBox( _
        "請輸入要篩選的欄號（數字，例如第2欄請輸入 2）：", "設定篩選欄位", "2")
    If colInput = "" Then Exit Sub
    If Not IsNumeric(colInput) Then
        MsgBox "請輸入有效欄號！", vbExclamation, "錯誤"
        Exit Sub
    End If
    colNum = CLng(colInput)
    If colNum < 1 Or colNum > lastCol Then
        MsgBox "欄號超出資料範圍（1 到 " & lastCol & "）！", vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 詢問前 N%
    userInput = InputBox( _
        "請輸入要篩選的前幾百分比（例如輸入 10 代表前 10%）：", "設定百分比", "10")
    If userInput = "" Then Exit Sub
    If Not IsNumeric(userInput) Then
        MsgBox "請輸入有效數字！", vbExclamation, "錯誤"
        Exit Sub
    End If
    pct = CDbl(userInput)
    If pct <= 0 Or pct > 100 Then
        MsgBox "百分比必須介於 0 到 100 之間！", vbExclamation, "錯誤"
        Exit Sub
    End If

    ' 計算門檻值（利用 PERCENTILE 工作表函數）
    Dim valRng As Range
    Set valRng = ws.Range(ws.Cells(2, colNum), ws.Cells(lastRow, colNum))
    threshold = Application.WorksheetFunction.Percentile(valRng, 1 - pct / 100)

    ' 建立篩選結果工作表
    pctInt = CLng(pct)
    newSheetName = "前" & pctInt & "pct篩選"
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(newSheetName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True

    Set newWs = ThisWorkbook.Sheets.Add(After:=ws)
    newWs.Name = newSheetName

    ' 複製標題列
    ws.Rows(1).Copy Destination:=newWs.Rows(1)
    destRow = 2

    Application.ScreenUpdating = False

    ' 篩選大於等於門檻值的資料列
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, colNum).Value) And Not IsEmpty(ws.Cells(i, colNum).Value) Then
            If CDbl(ws.Cells(i, colNum).Value) >= threshold Then
                ws.Rows(i).Copy Destination:=newWs.Rows(destRow)
                destRow = destRow + 1
            End If
        End If
    Next i

    newWs.Columns.AutoFit

    ' 加入篩選條件說明
    Dim infoRow As Long
    infoRow = destRow + 1
    newWs.Cells(infoRow, 1).Value = _
        "篩選條件：第 " & colNum & " 欄數值 >= " & threshold & _
        "（前 " & pct & "% 門檻值）"
    newWs.Cells(infoRow, 1).Font.Italic = True
    newWs.Cells(infoRow, 1).Font.Color = RGB(128, 128, 128)

    Application.ScreenUpdating = True

    MsgBox "前 " & pct & "% 篩選完成！" & vbNewLine & _
           "門檻值：" & threshold & vbNewLine & _
           "符合筆數：" & (destRow - 2) & " 筆", vbInformation, "完成"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
