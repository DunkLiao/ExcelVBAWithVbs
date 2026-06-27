Attribute VB_Name = "BatchDynamicArrayFormulas"
Option Explicit
'*************************************************************************************
'模組名稱: BatchDynamicArrayFormulas
'功能說明: 批次輸入動態陣列公式（UNIQUE、SORT、FILTER）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestBatchDynamicArrayFormulas()
    Call EnterBatchDynamicArrayFormulas("批次動態陣列公式")
End Sub

Sub EnterBatchDynamicArrayFormulas(ByVal sheetName As String)
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear

    ' 填入原始資料
    ws.Range("A1").Value = "產品編號"
    ws.Range("B1").Value = "產品名稱"
    ws.Range("C1").Value = "價格"
    ws.Range("D1").Value = "類別"
    ws.Range("A1:D1").Font.Bold = True

    ws.Range("A2").Value = "P001"
    ws.Range("B2").Value = "筆記型電腦"
    ws.Range("C2").Value = 35000
    ws.Range("D2").Value = "電子"

    ws.Range("A3").Value = "P002"
    ws.Range("B3").Value = "無線滑鼠"
    ws.Range("C3").Value = 1200
    ws.Range("D3").Value = "電子"

    ws.Range("A4").Value = "P003"
    ws.Range("B4").Value = "辦公椅"
    ws.Range("C4").Value = 4500
    ws.Range("D4").Value = "家具"

    ws.Range("A5").Value = "P001"
    ws.Range("B5").Value = "筆記型電腦"
    ws.Range("C5").Value = 35000
    ws.Range("D5").Value = "電子"

    ws.Range("A6").Value = "P004"
    ws.Range("B6").Value = "鍵盤"
    ws.Range("C6").Value = 800
    ws.Range("D6").Value = "電子"

    ws.Range("A7").Value = "P005"
    ws.Range("B7").Value = "書桌"
    ws.Range("C7").Value = 6200
    ws.Range("D7").Value = "家具"

    ws.Range("A8").Value = "P003"
    ws.Range("B8").Value = "辦公椅"
    ws.Range("C8").Value = 4500
    ws.Range("D8").Value = "家具"

    ' 標題列
    ws.Range("F1").Value = "UNIQUE 不重複產品"
    ws.Range("H1").Value = "SORT 價格排序"
    ws.Range("J1").Value = "FILTER 電子類別"
    ws.Range("F1").Font.Bold = True
    ws.Range("H1").Font.Bold = True
    ws.Range("J1").Font.Bold = True

    ' 批次輸入動態陣列公式
    ws.Range("F2").Formula2 = "=UNIQUE(A2:A8)"
    ws.Range("H2").Formula2 = "=SORT(B2:C8,2,-1)"
    ws.Range("J2").Formula2 = "=FILTER(A2:C8,D2:D8=""電子"")"

    ws.Columns("A:L").AutoFit

    MsgBox "批次動態陣列公式已輸入完成！" & vbCrLf & vbCrLf & _
           "F2: UNIQUE 列出不重複產品編號" & vbCrLf & _
           "H2: SORT 依價格降冪排序" & vbCrLf & _
           "J2: FILTER 篩選電子類別產品", vbInformation, "完成"
End Sub
