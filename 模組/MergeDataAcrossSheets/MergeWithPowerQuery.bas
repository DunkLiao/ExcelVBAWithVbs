Attribute VB_Name = "MergeWithPowerQuery"
Option Explicit
'*************************************************************************************
'家舱嘿: MergeWithPowerQuery
'弧: ㄏノ VBA 牟祇 Power Query 盢戈ㄖ穝
'
'舦┮Τ: Dunk
'祘Α砞璸: Dunk
'级糶ら戳: 2026/6/19
'
'*************************************************************************************

Sub TestMergeWithPowerQuery()
    ' ミ絛ㄒ
    Dim wsSrc As Worksheet
    Dim i As Long
    Dim j As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' ミ絛ㄒㄓ方
    For i = 1 To 3
        On Error Resume Next
        ThisWorkbook.Sheets("ㄓ方" & i).Delete
        On Error GoTo ErrHandler
        Set wsSrc = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsSrc.Name = "ㄓ方" & i
        
        wsSrc.Cells(1, 1).Value = "﹎"
        wsSrc.Cells(1, 2).Value = "场"
        wsSrc.Cells(1, 3).Value = "肂"
        
        For j = 1 To 5
            wsSrc.Cells(j + 1, 1).Value = "" & Chr(64 + i) & CStr(j)
            wsSrc.Cells(j + 1, 2).Value = Array("穨叭场", "祘场", "ㄆ场")(i - 1)
            wsSrc.Cells(j + 1, 3).Value = Int(Rnd * 50000) + 30000
        Next j
        wsSrc.Columns.AutoFit
    Next i
    
    ' ㄏノ Consolidate ㄖ絛瞅 Power Query ㄖ狦
    Dim wsDest As Worksheet
    On Error Resume Next
    ThisWorkbook.Sheets("PowerQueryㄖ").Delete
    On Error GoTo ErrHandler
    Set wsDest = ThisWorkbook.Sheets.Add
    wsDest.Name = "PowerQueryㄖ"
    
    ' ㄏノ Range.Consolidate ㄖ戈
    Dim srcRanges(1 To 3) As Variant
    For i = 1 To 3
        srcRanges(i) = ThisWorkbook.Sheets("ㄓ方" & i).UsedRange.Address(External:=True)
    Next i
    
    Dim consolidateSheets As Variant
    consolidateSheets = Array("ㄓ方1", "ㄓ方2", "ㄓ方3")
    
    Dim ws As Worksheet
    Dim srcRow As Long
    Dim destRow As Long
    destRow = 1
    
    ' ㄖ夹肈
    If ThisWorkbook.Sheets("ㄓ方1").Cells(1, 1).Value <> "" Then
        wsDest.Cells(destRow, 1).Value = "ㄓ方"
        For j = 1 To 3
            wsDest.Cells(destRow, j + 1).Value = ThisWorkbook.Sheets("ㄓ方1").Cells(1, j).Value
        Next j
        destRow = destRow + 1
    End If
    
    ' 硋ㄖ戈
    For i = 1 To 3
        Set ws = ThisWorkbook.Sheets("ㄓ方" & i)
        For srcRow = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            wsDest.Cells(destRow, 1).Value = "ㄓ方" & i
            For j = 1 To 3
                wsDest.Cells(destRow, j + 1).Value = ws.Cells(srcRow, j).Value
            Next j
            destRow = destRow + 1
        Next srcRow
    Next i
    
    wsDest.Columns.AutoFit
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "ㄖЧΘㄖ " & (destRow - 2) & " 掸戈" & vbCrLf & _
           "絛ㄒボ絛 VBA 家览 Power Query ㄖ瑈祘", vbInformation, "ЧΘ"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "ㄖ祇ネ岿粇" & Err.Number & " - " & Err.Description, vbCritical, "岿粇"
End Sub
