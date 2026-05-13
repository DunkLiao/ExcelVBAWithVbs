Attribute VB_Name = "NormalizeNumberFormat"
Option Explicit
'*************************************************************************************
'家舱嘿: NormalizeNumberFormat
'弧: 笆盢いゅ纗计锣传痷タ计
'          参甅ノだ┪计翴Α
'
'舦┮Τ: Dunk
'祘Α砞璸: Dunk
'级糶ら戳: 2026/5/13
'
'*************************************************************************************

Sub NormalizeNumberFormat()
    Dim ws           As Worksheet
    Dim rng          As Range
    Dim cell         As Range
    Dim convertCount As Long
    Dim skipCount    As Long
    Dim cleanVal     As String
    Dim numVal       As Double

    Set ws = ActiveSheet
    Set rng = ws.UsedRange

    convertCount = 0
    skipCount = 0

    Application.ScreenUpdating = False

    For Each cell In rng.Cells
        If cell.HasFormula = False And VarType(cell.Value) = vbString Then
            cleanVal = Trim(cell.Value)
            cleanVal = Replace(cleanVal, "$", "")
            cleanVal = Replace(cleanVal, "NT$", "")
            cleanVal = Replace(cleanVal, Chr(165), "")
            cleanVal = Replace(cleanVal, ",", "")
            cleanVal = Trim(cleanVal)

            If IsNumeric(cleanVal) Then
                numVal = CDbl(cleanVal)
                cell.Value = numVal
                If InStr(cleanVal, ".") > 0 Then
                    cell.NumberFormat = "#,##0.00"
                Else
                    cell.NumberFormat = "#,##0"
                End If
                convertCount = convertCount + 1
            Else
                skipCount = skipCount + 1
            End If
        End If
    Next cell

    Application.ScreenUpdating = True

    MsgBox "计Α夹非てЧΘ" & Chr(10) & _
        "锣传" & convertCount & " 纗" & Chr(10) & _
        "铬筁獶计" & skipCount & " 纗", _
        vbInformation, "ЧΘ"
End Sub

' ミゅ计代刚戈
Sub CreateNormalizeTestData()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("计Α代刚")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "计Α代刚"
    End If

    ws.Cells.Clear
    ws.Range("A1").Value = "﹍ゅ计"
    ws.Range("A2").Value = "1,234"
    ws.Range("A3").Value = "$5,678.90"
    ws.Range("A4").Value = "NT$9,000"
    ws.Range("A5").Value = "  42  "
    ws.Range("A6").Value = "3.14159"
    ws.Range("A7").Value = "獶计戈"
    ws.Range("A8").Value = "100,000"
    ws.Range("A1").Font.Bold = True
    ws.Columns.AutoFit

    MsgBox "代刚戈ミ叫癸 A2:A8 絛瞅磅︽ NormalizeNumberFormat", _
        vbInformation, "代刚戈ミЧΘ"
End Sub
