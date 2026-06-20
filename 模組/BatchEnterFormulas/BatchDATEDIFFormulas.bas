Attribute VB_Name = "BatchDATEDIFFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchDATEDIFFormulas
'功能說明: 批次輸入DATEDIF日期差異公式（年、月、日差異計算）的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestBatchDATEDIFFormulas()
    Call BatchEnterDATEDIFFormulas
End Sub

' 批次輸入DATEDIF公式
Sub BatchEnterDATEDIFFormulas()
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    
    sheetName = "DATEDIF公式"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillDATEDIFData(ws)
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 批次輸入DATEDIF公式
    For i = 2 To lastRow
        ' 計算年數差異
        ws.Cells(i, 4).Formula = "=DATEDIF(A" & i & ",B" & i & ",""Y"")"
        
        ' 計算總月數差異
        ws.Cells(i, 5).Formula = "=DATEDIF(A" & i & ",B" & i & ",""M"")"
        
        ' 計算月數（去除整年後剩餘月數）
        ws.Cells(i, 6).Formula = "=DATEDIF(A" & i & ",B" & i & ",""YM"")"
        
        ' 計算總天數差異
        ws.Cells(i, 7).Formula = "=DATEDIF(A" & i & ",B" & i & ",""D"")"
        
        ' 計算天數（去除整月後剩餘天數）
        ws.Cells(i, 8).Formula = "=DATEDIF(A" & i & ",B" & i & ",""MD"")"
        
        ' 組合文字格式的年月日差異
        ws.Cells(i, 9).Formula = _
            "=DATEDIF(A" & i & ",B" & i & ",""Y"")&""年""&" & _
            "DATEDIF(A" & i & ",B" & i & ",""YM"")&""個月""&" & _
            "DATEDIF(A" & i & ",B" & i & ",""MD"")&" & Chr(34) & "天" & Chr(34)
    Next i
    
    ws.Columns("A:I").AutoFit
    ws.Activate
    
    MsgBox "DATEDIF公式批次輸入完成！共 " & lastRow - 1 & " 筆。", vbInformation, "完成"
End Sub

' 填入DATEDIF示範資料
Private Sub FillDATEDIFData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "開始日期"
    ws.Range("B1").Value = "結束日期"
    ws.Range("C1").Value = "專案名稱"
    ws.Range("D1").Value = "年數(Y)"
    ws.Range("E1").Value = "總月數(M)"
    ws.Range("F1").Value = "剩餘月數(YM)"
    ws.Range("G1").Value = "總天數(D)"
    ws.Range("H1").Value = "剩餘天數(MD)"
    ws.Range("I1").Value = "年月日差異"
    
    ws.Range("A2").Value = "2020/1/15"
    ws.Range("B2").Value = "2024/6/1"
    ws.Range("C2").Value = "專案A"
    
    ws.Range("A3").Value = "2018/3/10"
    ws.Range("B3").Value = "2023/12/31"
    ws.Range("C3").Value = "專案B"
    
    ws.Range("A4").Value = "2021/7/1"
    ws.Range("B4").Value = "2025/3/15"
    ws.Range("C4").Value = "專案C"
    
    ws.Range("A5").Value = "2019/11/20"
    ws.Range("B5").Value = "2024/8/25"
    ws.Range("C5").Value = "專案D"
    
    ws.Range("A6").Value = "2022/1/5"
    ws.Range("B6").Value = "2025/1/5"
    ws.Range("C6").Value = "專案E"
    
    ws.Range("A7").Value = "2020/6/15"
    ws.Range("B7").Value = "2024/6/15"
    ws.Range("C7").Value = "專案F"
    
    ws.Columns("A:C").AutoFit
End Sub
