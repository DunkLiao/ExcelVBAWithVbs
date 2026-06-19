Attribute VB_Name = "CompareByFuzzyMatch"
Option Explicit
'*************************************************************************************
'模組名稱: CompareByFuzzyMatch
'功能說明: 使用模糊比對（Levenshtein 距離）比較兩個清單的相似度，標註差異
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestCompareByFuzzyMatch()
    Call CompareListsByFuzzyMatch
End Sub

Private Function LevenshteinDistance(ByVal s As String, ByVal t As String) As Long
    Dim d() As Long
    Dim m As Long, n As Long
    Dim i As Long, j As Long
    Dim cost As Long
    
    m = Len(s)
    n = Len(t)
    ReDim d(0 To m, 0 To n)
    
    For i = 0 To m
        d(i, 0) = i
    Next i
    For j = 0 To n
        d(0, j) = j
    Next j
    
    For i = 1 To m
        For j = 1 To n
            If Mid(s, i, 1) = Mid(t, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            d(i, j) = Application.Min(d(i - 1, j) + 1, d(i, j - 1) + 1, d(i - 1, j - 1) + cost)
        Next j
    Next i
    
    LevenshteinDistance = d(m, n)
End Function

Sub CompareListsByFuzzyMatch()
    Dim ws As Worksheet
    Dim lastRowA As Long
    Dim lastRowB As Long
    Dim i As Long
    Dim j As Long
    Dim bestMatch As String
    Dim bestDist As Long
    Dim currentDist As Long
    Dim strA As String
    Dim strB As String
    Dim destRow As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim wsName As String
    wsName = "模糊比對範例"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = wsName
    
    ' 清單 A（原始資料）
    ws.Range("A1").Value = "清單 A"
    ws.Range("B1").Value = "清單 B"
    ws.Range("D1").Value = "模糊比對結果"
    ws.Range("D1").Font.Bold = True
    ws.Range("D2").Value = "清單A名稱"
    ws.Range("E2").Value = "最佳匹配"
    ws.Range("F2").Value = "相似度"
    
    Dim listA As Variant
    listA = Array("台灣積體電路製造股份有限公司", "聯發科技股份有限公司", _
                   "鴻海精密工業股份有限公司", "台達電子工業股份有限公司", _
                   "中華電信股份有限公司")
    
    Dim listB As Variant
    listB = Array("台積電公司", "聯發科", "鴻海精密", _
                   "台達電子", "中華電信", "大立光電", "日月光半導體")
    
    For i = 0 To UBound(listA)
        ws.Cells(i + 2, 1).Value = listA(i)
    Next i
    
    For i = 0 To UBound(listB)
        ws.Cells(i + 2, 2).Value = listB(i)
    Next i
    
    lastRowA = UBound(listA) + 2
    lastRowB = UBound(listB) + 2
    
    ' 執行模糊比對
    destRow = 3
    For i = 1 To lastRowA - 1
        strA = ws.Cells(i + 1, 1).Value
        bestMatch = "(無匹配)"
        bestDist = 9999
        
        For j = 1 To lastRowB - 1
            strB = ws.Cells(j + 1, 2).Value
            currentDist = LevenshteinDistance(strA, strB)
            If currentDist < bestDist Then
                bestDist = currentDist
                bestMatch = strB
            End If
        Next j
        
        ws.Cells(destRow, 4).Value = strA
        ws.Cells(destRow, 5).Value = bestMatch
        
        Dim maxLen As Long
        maxLen = Application.Max(Len(strA), Len(bestMatch))
        If maxLen > 0 And bestMatch <> "(無匹配)" Then
            ws.Cells(destRow, 6).Value = Round((1 - bestDist / maxLen) * 100, 1) & "%"
        Else
            ws.Cells(destRow, 6).Value = "0%"
        End If
        
        destRow = destRow + 1
    Next i
    
    ws.Columns("A:B").AutoFit
    ws.Columns("D:F").AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "模糊比對完成！" & vbCrLf & _
           "共比對 " & (lastRowA - 1) & " 筆清單 A 與 " & (lastRowB - 1) & " 筆清單 B。" & vbCrLf & _
           "相似度使用 Levenshtein 距離演算法計算。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "模糊比對時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
