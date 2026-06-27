Attribute VB_Name = "CompareWithDrillDownDetail"
Option Explicit
'*************************************************************************************
'模組名稱: CompareWithDrillDownDetail
'功能說明: 自動比較新舊兩份資料差異，產生差異明細報表並標示調漲調降
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestCompareWithDrillDownDetail()
    Call CompareWithDrillDownDetail
End Sub

Sub CompareWithDrillDownDetail()
    Dim wsOld As Worksheet
    Dim wsNew As Worksheet
    Dim wsDiff As Worksheet
    Dim oldLastRow As Long
    Dim newLastRow As Long
    Dim diffRow As Long
    Dim i As Long
    Dim j As Long
    Dim oldKey As String
    Dim newKey As String
    Dim found As Boolean
    Dim oldVal As Double
    Dim newVal As Double
    
    On Error Resume Next
    Set wsOld = ThisWorkbook.Worksheets("舊版資料")
    On Error GoTo 0
    
    If wsOld Is Nothing Then
        Set wsOld = ThisWorkbook.Worksheets.Add
        wsOld.Name = "舊版資料"
    End If
    wsOld.Cells.Clear
    
    On Error Resume Next
    Set wsNew = ThisWorkbook.Worksheets("新版資料")
    On Error GoTo 0
    
    If wsNew Is Nothing Then
        Set wsNew = ThisWorkbook.Worksheets.Add
        wsNew.Name = "新版資料"
    End If
    wsNew.Cells.Clear
    
    Call FillCompareDrillData(wsOld, wsNew)
    
    On Error Resume Next
    Set wsDiff = ThisWorkbook.Worksheets("差異明細")
    If Not wsDiff Is Nothing Then wsDiff.Delete
    On Error GoTo 0
    
    Set wsDiff = ThisWorkbook.Worksheets.Add
    wsDiff.Name = "差異明細"
    
    wsDiff.Range("A1").Value = "產品編號"
    wsDiff.Range("B1").Value = "產品名稱"
    wsDiff.Range("C1").Value = "舊版價格"
    wsDiff.Range("D1").Value = "新版價格"
    wsDiff.Range("E1").Value = "差異金額"
    wsDiff.Range("F1").Value = "差異百分比"
    wsDiff.Range("G1").Value = "類型"
    
    oldLastRow = wsOld.Cells(wsOld.Rows.Count, 1).End(xlUp).Row
    newLastRow = wsNew.Cells(wsNew.Rows.Count, 1).End(xlUp).Row
    diffRow = 1
    
    For i = 2 To oldLastRow
        oldKey = CStr(wsOld.Cells(i, 1).Value)
        found = False
        
        For j = 2 To newLastRow
            newKey = CStr(wsNew.Cells(j, 1).Value)
            
            If oldKey = newKey Then
                oldVal = CDbl(wsOld.Cells(i, 3).Value)
                newVal = CDbl(wsNew.Cells(j, 3).Value)
                
                If oldVal <> newVal Then
                    diffRow = diffRow + 1
                    wsDiff.Cells(diffRow, 1).Value = oldKey
                    wsDiff.Cells(diffRow, 2).Value = wsOld.Cells(i, 2).Value
                    wsDiff.Cells(diffRow, 3).Value = oldVal
                    wsDiff.Cells(diffRow, 4).Value = newVal
                    wsDiff.Cells(diffRow, 5).Formula = "=D" & diffRow & "-C" & diffRow
                    
                    If oldVal <> 0 Then
                        wsDiff.Cells(diffRow, 6).Formula = "=(D" & diffRow & "-C" & diffRow & ")/C" & diffRow
                        wsDiff.Cells(diffRow, 6).NumberFormat = "0.00%"
                    End If
                    
                    If newVal > oldVal Then
                        wsDiff.Cells(diffRow, 7).Value = "調漲"
                    Else
                        wsDiff.Cells(diffRow, 7).Value = "調降"
                    End If
                End If
                found = True
                Exit For
            End If
        Next j
        
        If Not found Then
            diffRow = diffRow + 1
            wsDiff.Cells(diffRow, 1).Value = oldKey
            wsDiff.Cells(diffRow, 2).Value = wsOld.Cells(i, 2).Value
            wsDiff.Cells(diffRow, 3).Value = wsOld.Cells(i, 3).Value
            wsDiff.Cells(diffRow, 4).Value = "已移除"
            wsDiff.Cells(diffRow, 7).Value = "下架"
        End If
    Next i
    
    With wsDiff.Range("E2:F" & diffRow)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Interior.Color = RGB(255, 200, 200)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(2).Interior.Color = RGB(200, 255, 200)
    End With
    
    wsDiff.Columns("A:G").AutoFit
    wsDiff.Activate
    
    MsgBox "差異比較完成！共發現 " & diffRow - 1 & " 筆差異。", vbInformation, "完成"
End Sub

Private Sub FillCompareDrillData(ByVal wsOld As Worksheet, ByVal wsNew As Worksheet)
    wsOld.Range("A1").Value = "編號"
    wsOld.Range("B1").Value = "品名"
    wsOld.Range("C1").Value = "價格"
    wsOld.Range("A2").Value = "P01"
    wsOld.Range("B2").Value = "原子筆"
    wsOld.Range("C2").Value = 25
    wsOld.Range("A3").Value = "P02"
    wsOld.Range("B3").Value = "筆記本"
    wsOld.Range("C3").Value = 60
    wsOld.Range("A4").Value = "P03"
    wsOld.Range("B4").Value = "資料夾"
    wsOld.Range("C4").Value = 35
    wsOld.Range("A5").Value = "P04"
    wsOld.Range("B5").Value = "膠帶"
    wsOld.Range("C5").Value = 15
    wsOld.Range("A6").Value = "P05"
    wsOld.Range("B6").Value = "剪刀"
    wsOld.Range("C6").Value = 80
    
    wsNew.Range("A1").Value = "編號"
    wsNew.Range("B1").Value = "品名"
    wsNew.Range("C1").Value = "價格"
    wsNew.Range("A2").Value = "P01"
    wsNew.Range("B2").Value = "原子筆"
    wsNew.Range("C2").Value = 30
    wsNew.Range("A3").Value = "P02"
    wsNew.Range("B3").Value = "筆記本"
    wsNew.Range("C3").Value = 60
    wsNew.Range("A4").Value = "P03"
    wsNew.Range("B4").Value = "資料夾"
    wsNew.Range("C4").Value = 40
    wsNew.Range("A5").Value = "P04"
    wsNew.Range("B5").Value = "膠帶"
    wsNew.Range("C5").Value = 12
    wsNew.Range("A6").Value = "P06"
    wsNew.Range("B6").Value = "尺"
    wsNew.Range("C6").Value = 25
    
    wsOld.Columns("A:C").AutoFit
    wsNew.Columns("A:C").AutoFit
End Sub
