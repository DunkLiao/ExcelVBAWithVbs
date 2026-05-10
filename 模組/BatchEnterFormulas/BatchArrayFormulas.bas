Attribute VB_Name = "BatchArrayFormulas"
Option Explicit

' ============================================================
' 模組名稱：BatchArrayFormulas
' 功能說明：批次在多個欄位輸入陣列公式範例
'           包含：SUMPRODUCT、多條件加總、矩陣乘法等
' 注意：陣列公式需以 Ctrl+Shift+Enter 輸入（或使用 VBA FormulaArray）
' ============================================================

Sub CreateBatchArrayFormulas()
    Dim ws      As Worksheet
    Dim wsName  As String
    Dim i       As Long
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    wsName = "陣列公式範例"
    
    ' 刪除舊工作表
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets(wsName).Delete
    On Error GoTo ErrHandler
    Application.DisplayAlerts = True
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = wsName
    
    ' === 建立原始資料區 ===
    ws.Range("A1:E1").Value = Array("產品", "數量", "單價", "類別", "折扣率")
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(68, 114, 196)
    ws.Rows(1).Font.Color = RGB(255, 255, 255)
    
    ' 填入範例資料
    Dim products(9) As String
    Dim qtys(9)     As Integer
    Dim prices(9)   As Double
    Dim cats(9)     As String
    Dim discounts(9) As Double
    
    products(0) = "蘋果" : qtys(0) = 100 : prices(0) = 25.0 : cats(0) = "水果" : discounts(0) = 0.05
    products(1) = "香蕉" : qtys(1) = 200 : prices(1) = 12.0 : cats(1) = "水果" : discounts(1) = 0.0
    products(2) = "牛奶" : qtys(2) = 50  : prices(2) = 65.0 : cats(2) = "乳製品" : discounts(2) = 0.1
    products(3) = "起司" : qtys(3) = 30  : prices(3) = 120.0: cats(3) = "乳製品" : discounts(3) = 0.05
    products(4) = "雞蛋" : qtys(4) = 80  : prices(4) = 45.0 : cats(4) = "蛋品"   : discounts(4) = 0.0
    products(5) = "葡萄" : qtys(5) = 60  : prices(5) = 80.0 : cats(5) = "水果" : discounts(5) = 0.08
    products(6) = "優格" : qtys(6) = 40  : prices(6) = 55.0 : cats(6) = "乳製品" : discounts(6) = 0.05
    products(7) = "橘子" : qtys(7) = 150 : prices(7) = 18.0 : cats(7) = "水果" : discounts(7) = 0.0
    products(8) = "鮮奶油": qtys(8) = 20 : prices(8) = 95.0 : cats(8) = "乳製品" : discounts(8) = 0.1
    products(9) = "鵪鶉蛋": qtys(9) = 60 : prices(9) = 35.0 : cats(9) = "蛋品"   : discounts(9) = 0.0
    
    For i = 0 To 9
        ws.Cells(i + 2, 1).Value = products(i)
        ws.Cells(i + 2, 2).Value = qtys(i)
        ws.Cells(i + 2, 3).Value = prices(i)
        ws.Cells(i + 2, 4).Value = cats(i)
        ws.Cells(i + 2, 5).Value = discounts(i)
    Next i
    
    ' === 公式結果區（G欄起） ===
    ws.Range("G1").Value = "公式說明"
    ws.Range("H1").Value = "計算結果"
    ws.Rows(1).Font.Bold = True
    
    ' --- 公式 1：SUMPRODUCT 計算加權總金額（數量×單價）---
    ws.Range("G3").Value = "SUMPRODUCT 數量×單價總金額"
    ws.Range("H3").FormulaArray = "=SUMPRODUCT(B2:B11,C2:C11)"
    ws.Range("H3").NumberFormat = "#,##0.00"
    
    ' --- 公式 2：多條件陣列加總（水果類別的總金額）---
    ws.Range("G4").Value = "陣列加總：水果類數量×單價"
    ws.Range("H4").FormulaArray = "=SUM(IF(D2:D11=""水果"",B2:B11*C2:C11,0))"
    ws.Range("H4").NumberFormat = "#,##0.00"
    
    ' --- 公式 3：扣除折扣後的總金額（SUMPRODUCT）---
    ws.Range("G5").Value = "SUMPRODUCT 含折扣後總金額"
    ws.Range("H5").FormulaArray = "=SUMPRODUCT(B2:B11,C2:C11,(1-E2:E11))"
    ws.Range("H5").NumberFormat = "#,##0.00"
    
    ' --- 公式 4：陣列計算各類別數量（水果）---
    ws.Range("G6").Value = "陣列統計：水果類商品種數"
    ws.Range("H6").FormulaArray = "=SUM((D2:D11=""水果"")*1)"
    
    ' --- 公式 5：陣列找最大折扣金額 ---
    ws.Range("G7").Value = "陣列：最大折扣金額"
    ws.Range("H7").FormulaArray = "=MAX(B2:B11*C2:C11*E2:E11)"
    ws.Range("H7").NumberFormat = "#,##0.00"
    
    ' --- 公式 6：SUMPRODUCT 多條件（乳製品且折扣>0）---
    ws.Range("G8").Value = "SUMPRODUCT 乳製品且有折扣的金額"
    ws.Range("H8").Formula = "=SUMPRODUCT((D2:D11=""乳製品"")*(E2:E11>0)*B2:B11*C2:C11)"
    ws.Range("H8").NumberFormat = "#,##0.00"
    
    ' --- 公式 7：陣列計算平均單價（排除折扣=0的品項）---
    ws.Range("G9").Value = "陣列：有折扣品項的平均單價"
    ws.Range("H9").FormulaArray = "=AVERAGE(IF(E2:E11>0,C2:C11))"
    ws.Range("H9").NumberFormat = "#,##0.00"
    
    ' 設定樣式
    With ws.Range("G3:G9")
        .Interior.Color = RGB(242, 242, 242)
        .WrapText = False
    End With
    ws.Columns("A:H").AutoFit
    ws.Range("G:G").ColumnWidth = 35
    
    Application.ScreenUpdating = True
    MsgBox "陣列公式範例建立完成！" & vbCrLf & _
           "請查看「" & wsName & "」工作表。", vbInformation, "完成"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub