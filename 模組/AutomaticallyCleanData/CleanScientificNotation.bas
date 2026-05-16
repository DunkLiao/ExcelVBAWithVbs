Attribute VB_Name = "CleanScientificNotation"
Option Explicit
'*************************************************************************************
'模組名稱: CleanScientificNotation
'功能說明: 掃描指定欄位，將以科學記號格式（如 1.23E+05）儲存的數值
'          轉換為一般整數或小數格式，並選擇性地套用千位分隔符號
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

' 清理作用中工作表選取範圍的科學記號
Sub CleanScientificNotationInSelection()
    Dim rng     As Range
    Dim cell    As Range
    Dim strVal  As String
    Dim dblVal  As Double
    Dim count   As Long

    If TypeName(Selection) <> "Range" Then
        MsgBox "請先選取要清理的儲存格範圍。", vbExclamation
        Exit Sub
    End If

    Set rng = Selection
    count = 0

    Application.ScreenUpdating = False

    For Each cell In rng.Cells
        If cell.Value <> "" Then
            strVal = cell.Text  ' 取得顯示文字（含科學記號）

            ' 判斷是否為科學記號格式
            If InStr(1, UCase(strVal), "E+") > 0 Or _
               InStr(1, UCase(strVal), "E-") > 0 Then
                If IsNumeric(cell.Value) Then
                    dblVal = CDbl(cell.Value)
                    cell.Value = dblVal

                    ' 判斷是否為整數，設定格式
                    If dblVal = CLng(dblVal) Then
                        cell.NumberFormat = "#,##0"
                    Else
                        cell.NumberFormat = "#,##0.00"
                    End If
                    count = count + 1
                End If
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    If count = 0 Then
        MsgBox "選取範圍中未發現科學記號格式的數值。", vbInformation, "完成"
    Else
        MsgBox "已清理 " & count & " 個科學記號數值。", vbInformation, "完成"
    End If
End Sub

' 清理整欄的科學記號（依欄位代號）
Sub CleanScientificNotationInColumn()
    Dim ws       As Worksheet
    Dim colInput As String
    Dim colIdx   As Long
    Dim lastRow  As Long
    Dim r        As Long
    Dim dblVal   As Double
    Dim strVal   As String
    Dim count    As Long

    Set ws = ActiveSheet

    colInput = InputBox("請輸入要清理的欄位代號（例如 A）：", "清理科學記號", "A")
    If colInput = "" Then Exit Sub

    On Error Resume Next
    colIdx = ws.Range(colInput & "1").Column
    On Error GoTo 0

    lastRow = ws.Cells(ws.Rows.Count, colIdx).End(xlUp).Row
    count = 0

    Application.ScreenUpdating = False

    For r = 1 To lastRow
        If ws.Cells(r, colIdx).Value <> "" Then
            strVal = ws.Cells(r, colIdx).Text
            If InStr(1, UCase(strVal), "E+") > 0 Or _
               InStr(1, UCase(strVal), "E-") > 0 Then
                If IsNumeric(ws.Cells(r, colIdx).Value) Then
                    dblVal = CDbl(ws.Cells(r, colIdx).Value)
                    ws.Cells(r, colIdx).Value = dblVal
                    If dblVal = CLng(dblVal) Then
                        ws.Cells(r, colIdx).NumberFormat = "#,##0"
                    Else
                        ws.Cells(r, colIdx).NumberFormat = "#,##0.00"
                    End If
                    count = count + 1
                End If
            End If
        End If
    Next r

    Application.ScreenUpdating = True
    MsgBox "欄位 " & colInput & " 共清理 " & count & " 個科學記號數值。", vbInformation, "完成"
End Sub

' 建立示範資料並執行清理
Sub DemoCleanScientificNotation()
    Dim ws As Worksheet
    Set ws = GetOrCreateSciSheet("科學記號清理示範")
    ws.Cells.Clear

    ws.Range("A1:C1").Value = Array("品項", "數量（科學記號）", "清理後")
    ws.Range("A1:C1").Font.Bold = True

    ' 強制以科學記號格式輸入數值
    ws.Range("B2").Value = 123456
    ws.Range("B2").NumberFormat = "0.00E+00"
    ws.Range("B3").Value = 9870000
    ws.Range("B3").NumberFormat = "0.00E+00"
    ws.Range("B4").Value = 0.000456
    ws.Range("B4").NumberFormat = "0.00E+00"
    ws.Range("B5").Value = 1500000000
    ws.Range("B5").NumberFormat = "0.00E+00"

    ws.Range("A2").Value = "螺絲"
    ws.Range("A3").Value = "電阻"
    ws.Range("A4").Value = "電容值"
    ws.Range("A5").Value = "資料筆數"

    ' 複製到C欄並清理
    ws.Range("B2:B5").Copy ws.Range("C2")

    Dim r As Long
    Dim dblVal As Double
    For r = 2 To 5
        If IsNumeric(ws.Cells(r, 3).Value) Then
            dblVal = CDbl(ws.Cells(r, 3).Value)
            ws.Cells(r, 3).Value = dblVal
            If Abs(dblVal) >= 1 And dblVal = CLng(dblVal) Then
                ws.Cells(r, 3).NumberFormat = "#,##0"
            Else
                ws.Cells(r, 3).NumberFormat = "#,##0.000000"
            End If
        End If
    Next r

    ws.Columns("A:C").AutoFit
    MsgBox "科學記號清理示範完成！對比 B 欄（科學記號）與 C 欄（清理後）。", _
           vbInformation, "完成"
End Sub

Private Function GetOrCreateSciSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateSciSheet = ws
End Function
