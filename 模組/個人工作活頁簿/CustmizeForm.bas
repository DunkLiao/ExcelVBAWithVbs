Attribute VB_Name = "CustmizeForm"
Sub 建立問題追蹤表()
    Application.ScreenUpdating = False
    
    Dim tableName As String
    tableName = "問題追蹤表_" & ActiveSheet.Name
    
    ActiveSheet.Name = tableName
    ActiveCell.FormulaR1C1 = "單號"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "優先程度"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "紀錄日期"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "預計完成日期"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "實際完成日期"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "處理結果"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "聯絡窗口"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "結案日期"
    Range("A1").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$H$1"), , xlYes).Name = _
        tableName
    Range(tableName & "[#All]").Select
    ActiveSheet.ListObjects(tableName).TableStyle = "TableStyleLight11"
    Range(tableName & "[紀錄日期]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormatLocal = "[$-404]e/m/d;@"
    ActiveWindow.SmallScroll Down:=-12
    Range(tableName & "[預計完成日期]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormatLocal = "[$-404]e/m/d;@"
    ActiveWindow.SmallScroll Down:=-6
    Range(tableName & "[實際完成日期]").Select
    ActiveWindow.SmallScroll Down:=-6
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormatLocal = "[$-404]e/m/d;@"
    ActiveWindow.SmallScroll Down:=-12
    Range(tableName & "[結案日期]").Select
    ActiveWindow.SmallScroll Down:=-18
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormatLocal = "[$-404]e/m/d;@"
    ActiveWindow.SmallScroll Down:=-12
    Range("A1").Select
    
    畫所有格線 (tableName)
    
    Application.ScreenUpdating = True
    
    
End Sub

Sub 建立資料庫表格定義()
    Application.ScreenUpdating = False
    
    Dim tableName As String
    tableName = "table_" & ActiveSheet.Name
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "欄位名稱"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "欄位描述"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "型態"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "長度"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "主鍵(PK)"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "說明"
    Range("F1").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$F$1"), , xlYes).Name = _
        tableName
    ActiveSheet.ListObjects(tableName).TableStyle = "TableStyleLight11"
    
    '建立回總表連結
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "回總表"
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "總表索引!A1", TextToDisplay:="回總表"
    
    畫所有格線 (tableName)
    
    Application.ScreenUpdating = True
    
End Sub


Sub 建立資料庫視圖定義()
    Application.ScreenUpdating = False

    Dim tableName As String
    tableName = "view_" & ActiveSheet.Name
    
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "視圖名稱"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "視圖定義(create sql)"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "說明"
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$C$1"), , xlYes).Name = _
        tableName
    ActiveSheet.ListObjects(tableName).TableStyle = "TableStyleLight11"
    
    '建立回總表連結
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "回總表"
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "總表索引!A1", TextToDisplay:="回總表"
        
    畫所有格線 (tableName)
    
    Application.ScreenUpdating = True
        
End Sub


Sub 畫所有格線(ByVal objName As String)
Application.Goto Reference:=objName
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub


