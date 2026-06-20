Attribute VB_Name = "MergeWithFilterAndSort"
Option Explicit
'*************************************************************************************
'模組名稱: MergeWithFilterAndSort
'功能說明: 合併跨工作表資料，並自動篩選非空白及排序的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestMergeWithFilterAndSort()
    Call MergeWithFilterAndSort
End Sub

' 合併所有工作表資料，篩選指定欄位非空白，並依金額欄排序
Sub MergeWithFilterAndSort()
    Dim wsTarget As Worksheet
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim sourceLastRow As Long
    Dim i As Long
    
    ' 建立目標工作表
    On Error Resume Next
    ThisWorkbook.Worksheets("合併篩選排序").Delete
    On Error GoTo 0
    
    Set wsTarget = ThisWorkbook.Worksheets.Add
    wsTarget.Name = "合併篩選排序"
    wsTarget.Range("A1").Value = "產品名稱"
    wsTarget.Range("B1").Value = "分類"
    wsTarget.Range("C1").Value = "金額"
    wsTarget.Range("D1").Value = "來源工作表"
    
    targetRow = 1
    
    ' 遍歷所有工作表
    For Each ws In ThisWorkbook.Worksheets
        ' 跳過目標工作表
        If ws.Name <> "合併篩選排序" Then
            sourceLastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            If sourceLastRow > 1 Then
                ' 複製資料（略過標題列）
                ws.Range("A2:C" & sourceLastRow).Copy
                wsTarget.Cells(targetRow + 1, 1).PasteSpecial xlPasteValues
                
                ' 填入來源工作表名稱
                wsTarget.Range("D" & targetRow + 1 & ":D" & targetRow + sourceLastRow - 1).Value = ws.Name
            End If
            
            targetRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
        End If
    Next ws
    
    ' 刪除C欄為空白的列（篩選功能）
    If targetRow > 1 Then
        For i = targetRow To 2 Step -1
            If IsEmpty(wsTarget.Cells(i, 3)) Or Len(CStr(wsTarget.Cells(i, 3).Value)) = 0 Then
                wsTarget.Rows(i).Delete
            End If
        Next i
    End If
    
    ' 重新取得最後列
    targetRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
    
    ' 依金額欄排序（降冪）
    If targetRow > 1 Then
        With wsTarget.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsTarget.Range("C2:C" & targetRow), _
                SortOn:=xlSortOnValues, _
                Order:=xlDescending
            .SetRange wsTarget.Range("A1:D" & targetRow)
            .Header = xlYes
            .Apply
        End With
    End If
    
    wsTarget.Columns("A:D").AutoFit
    wsTarget.Activate
    
    MsgBox "跨工作表資料已合併、篩選並排序完成！", vbInformation, "完成"
End Sub
