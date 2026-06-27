Attribute VB_Name = "MergeSheetsByCustomGroupKey"
Option Explicit
'*************************************************************************************
'模組名稱: MergeSheetsByCustomGroupKey
'功能說明: 以自訂群組鍵合併跨工作表資料，自動按月份分組彙總的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestMergeSheetsByCustomGroupKey()
    Call MergeSheetsByCustomGroupKey
End Sub

Sub MergeSheetsByCustomGroupKey()
    Dim wsTarget As Worksheet
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim sourceLastRow As Long
    Dim i As Long
    Dim monthKey As String
    Dim groupDict As Object
    Dim keyVal As Variant
    
    On Error Resume Next
    ThisWorkbook.Worksheets("群組合併結果").Delete
    On Error GoTo 0
    
    Set wsTarget = ThisWorkbook.Worksheets.Add
    wsTarget.Name = "群組合併結果"
    wsTarget.Range("A1").Value = "月份"
    wsTarget.Range("B1").Value = "合計金額"
    wsTarget.Range("C1").Value = "筆數"
    
    Set groupDict = CreateObject("Scripting.Dictionary")
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "群組合併結果" Then
            sourceLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            If sourceLastRow > 1 Then
                For i = 2 To sourceLastRow
                    monthKey = CStr(ws.Cells(i, 1).Value)
                    If Len(monthKey) > 0 Then
                        If Not groupDict.Exists(monthKey) Then
                            groupDict.Add monthKey, Array(0, 0)
                        End If
                        Dim arr
                        arr = groupDict(monthKey)
                        arr(0) = arr(0) + CLng(ws.Cells(i, 2).Value)
                        arr(1) = arr(1) + 1
                        groupDict(monthKey) = arr
                    End If
                Next i
            End If
        End If
    Next ws
    
    targetRow = 1
    For Each keyVal In groupDict.Keys
        targetRow = targetRow + 1
        Dim vals
        vals = groupDict(keyVal)
        wsTarget.Cells(targetRow, 1).Value = keyVal
        wsTarget.Cells(targetRow, 2).Value = vals(0)
        wsTarget.Cells(targetRow, 3).Value = vals(1)
    Next keyVal
    
    If targetRow > 1 Then
        With wsTarget.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsTarget.Range("A2:A" & targetRow), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending
            .SetRange wsTarget.Range("A1:C" & targetRow)
            .Header = xlYes
            .Apply
        End With
    End If
    
    wsTarget.Columns("A:C").AutoFit
    wsTarget.Activate
    Set groupDict = Nothing
    
    MsgBox "跨工作表群組合併完成！共 " & targetRow - 1 & " 個群組。", vbInformation, "完成"
End Sub
