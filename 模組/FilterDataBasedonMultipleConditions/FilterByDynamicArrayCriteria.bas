Attribute VB_Name = "FilterByDynamicArrayCriteria"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByDynamicArrayCriteria
'功能說明: 使用動態陣列公式產生的條件來篩選資料的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestFilterByDynamicArrayCriteria()
    Call DynamicArrayCriteriaFilter("動態陣列篩選範例")
End Sub

Sub DynamicArrayCriteriaFilter(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim criteriaRange As Range
    Dim dataRange As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear

    ' 填入主資料
    ws.Range("A1").Value = "訂單編號"
    ws.Range("B1").Value = "客戶名稱"
    ws.Range("C1").Value = "產品"
    ws.Range("D1").Value = "金額"
    ws.Range("E1").Value = "日期"
    ws.Range("A1:E1").Font.Bold = True

    ws.Range("A2").Value = "ORD001"
    ws.Range("B2").Value = "甲公司"
    ws.Range("C2").Value = "產品A"
    ws.Range("D2").Value = 25000
    ws.Range("E2").Value = "2026/1/5"

    ws.Range("A3").Value = "ORD002"
    ws.Range("B3").Value = "乙公司"
    ws.Range("C3").Value = "產品B"
    ws.Range("D3").Value = 18000
    ws.Range("E3").Value = "2026/1/8"

    ws.Range("A4").Value = "ORD003"
    ws.Range("B4").Value = "甲公司"
    ws.Range("C4").Value = "產品A"
    ws.Range("D4").Value = 32000
    ws.Range("E4").Value = "2026/1/15"

    ws.Range("A5").Value = "ORD004"
    ws.Range("B5").Value = "丙公司"
    ws.Range("C5").Value = "產品C"
    ws.Range("D5").Value = 15000
    ws.Range("E5").Value = "2026/1/20"

    ws.Range("A6").Value = "ORD005"
    ws.Range("B6").Value = "乙公司"
    ws.Range("C6").Value = "產品A"
    ws.Range("D6").Value = 28000
    ws.Range("E6").Value = "2026/1/22"

    ws.Range("A7").Value = "ORD006"
    ws.Range("B7").Value = "甲公司"
    ws.Range("C7").Value = "產品B"
    ws.Range("D7").Value = 12000
    ws.Range("E7").Value = "2026/1/25"

    ws.Range("A8").Value = "ORD007"
    ws.Range("B8").Value = "丁公司"
    ws.Range("C8").Value = "產品A"
    ws.Range("D7").Value = 22000
    ws.Range("E7").Value = "2026/1/28"

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' 使用動態陣列公式建立篩選條件（UNIQUE 列出不重複客戶）
    ws.Range("G1").Value = "不重複客戶（動態陣列）"
    ws.Range("G1").Font.Bold = True
    ws.Range("G2").Formula2 = "=UNIQUE(B2:B" & lastRow & ")"

    ' 使用動態陣列篩選：SORT 將客戶依金額排序
    ws.Range("I1").Value = "高金額客戶（動態篩選）"
    ws.Range("I1").Font.Bold = True
    ws.Range("I2").Formula2 = "=FILTER(A2:E" & lastRow & ",D2:D" & lastRow & ">20000)"


    ' 使用 VBA 讀取動態陣列的溢出範圍，並對原始資料進行自動篩選
    Dim criteriaRng As Range
    On Error Resume Next
    Set criteriaRng = ws.Range("G2").CurrentRegion
    On Error GoTo 0

    ' 使用進階篩選：將原始資料複製到新區域，以 J 欄開始
    ws.Range("J1").Value = "VBA進階篩選結果"
    ws.Range("J1").Font.Bold = True

    Set dataRange = ws.Range("A1:E" & lastRow)

    ' 建立準則範圍：篩選金額 > 20000
    ws.Range("H10").Value = "金額"
    ws.Range("H11").Value = ">20000"

    dataRange.AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=ws.Range("H10:H11"), _
        CopyToRange:=ws.Range("J2"), _
        Unique:=False

    ' 清除暫存準則
    ws.Range("H10:H11").Clear

    ws.Columns("A:L").AutoFit

    MsgBox "動態陣列篩選完成！" & vbCrLf & vbCrLf & _
           "G2: UNIQUE 公式列出不重複客戶" & vbCrLf & _
           "I2: FILTER 公式動態篩選高金額訂單" & vbCrLf & _
           "J2: VBA 進階篩選結果", vbInformation, "完成"
End Sub
