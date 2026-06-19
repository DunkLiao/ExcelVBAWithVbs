Attribute VB_Name = "FilterByWildcardPattern"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByWildcardPattern
'功能說明: 使用萬用字元（* 和 ?）模式比對篩選資料，支援多個模式同時篩選
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

Sub TestFilterByWildcardPattern()
    Call FilterByWildcard
End Sub

Sub FilterByWildcard()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rngData As Range
    Dim i As Long
    Dim pattern As String
    Dim filterHeader As String
    
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    Dim wsName As String
    wsName = "萬用字元篩選"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler
    
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = wsName
    
    ' 建立範例資料
    ws.Range("A1").Value = "產品編號"
    ws.Range("B1").Value = "產品名稱"
    ws.Range("C1").Value = "供應商"
    ws.Range("D1").Value = "庫存量"
    
    Dim products As Variant
    products = Array( _
        Array("NB-ASUS-001", "ASUS 筆記型電腦 Pro", "華碩電腦", 150), _
        Array("NB-ACER-002", "Acer 筆記型電腦 Lite", "宏碁股份", 80), _
        Array("TB-APPLE-001", "Apple iPad Pro 12.9", "蘋果亞洲", 200), _
        Array("TB-SAMS-001", "Samsung Galaxy Tab", "三星電子", 120), _
        Array("PH-APPLE-001", "Apple iPhone 15", "蘋果亞洲", 300), _
        Array("PH-SAMS-001", "Samsung Galaxy S24", "三星電子", 250), _
        Array("NB-DELL-001", "Dell 筆記型電腦 XPS", "戴爾電腦", 60), _
        Array("TB-ASUS-001", "ASUS ZenPad 10", "華碩電腦", 90), _
        Array("AC-BOSE-001", "Bose 降噪耳機 QC45", "博士音響", 45), _
        Array("AC-SONY-001", "Sony 無線耳機 WH-1000", "索尼公司", 70))
    
    For i = 0 To 9
        ws.Cells(i + 2, 1).Value = products(i)(0)
        ws.Cells(i + 2, 2).Value = products(i)(1)
        ws.Cells(i + 2, 3).Value = products(i)(2)
        ws.Cells(i + 2, 4).Value = products(i)(3)
    Next i
    
    lastRow = 11
    Set rngData = ws.Range("A1:D" & lastRow)
    
    ' 提示使用者輸入萬用字元模式
    pattern = InputBox("請輸入萬用字元篩選模式（例如 *APPLE*）：" & vbCrLf & _
                        "萬用字元說明：* 代表任意字串，? 代表單一字元", _
                        "萬用字元篩選", "*APPLE*")
    If pattern = "" Then
        MsgBox "已取消操作", vbInformation, "取消"
        GoTo CleanUp
    End If
    
    ' 使用 AutoFilter 搭配萬用字元篩選
    filterHeader = "產品編號"
    
    rngData.AutoFilter Field:=1, Criteria1:="=" & pattern, Operator:=xlOr, Criteria2:="=" & pattern
    rngData.AutoFilter Field:=2, Criteria1:="=" & pattern, Operator:=xlOr, Criteria2:="=" & pattern
    
    ' 將可見的篩選結果複製到下方
    Dim visibleRange As Range
    Dim destCell As Range
    Dim rowCount As Long
    
    On Error Resume Next
    Set visibleRange = rngData.Offset(1, 0).Resize(rngData.Rows.Count - 1).SpecialCells(xlCellTypeVisible)
    On Error GoTo ErrHandler
    
    If Not visibleRange Is Nothing Then
        Set destCell = ws.Range("A" & (lastRow + 3))
        destCell.Value = "篩選結果（模式：" & pattern & "）"
        destCell.Font.Bold = True
        
        ' 複製標題
        rngData.Rows(1).Copy ws.Range("A" & (lastRow + 4))
        visibleRange.Copy ws.Range("A" & (lastRow + 5))
        
        rowCount = visibleRange.Rows.Count
    Else
        ws.Range("A" & (lastRow + 3)).Value = "沒有符合模式 " & pattern & " 的資料"
        rowCount = 0
    End If
    
    ' 取消篩選
    rngData.AutoFilter
    
    ws.Columns.AutoFit
    
    MsgBox "萬用字元篩選完成！" & vbCrLf & _
           "篩選模式：" & pattern & vbCrLf & _
           "符合筆數：" & rowCount & " 筆", vbInformation, "完成"
    
CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
    
ErrHandler:
    On Error Resume Next
    rngData.AutoFilter
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "篩選時發生錯誤：" & Err.Number & " - " & Err.Description, vbCritical, "錯誤"
End Sub
