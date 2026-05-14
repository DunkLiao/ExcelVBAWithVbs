Attribute VB_Name = "CubeFormulaExample"
Option Explicit
'*************************************************************************************
'模組名稱: CubeFormulaExample
'功能說明: 示範以 VBA 建立 CUBEVALUE、CUBEMEMBER 等 Cube 函數說明的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點
Sub TestCubeFormula()
    Call CreateCubeFormulaSheet("Cube函數範例")
End Sub

' 建立 Cube 函數示範工作表
' sheetName: 目標工作表名稱
Sub CreateCubeFormulaSheet(ByVal sheetName As String)
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateCubeSheet(sheetName)
    ws.Cells.Clear

    ws.Range("A1").Value = "函數名稱"
    ws.Range("B1").Value = "公式字串（示意）"
    ws.Range("C1").Value = "說明"

    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    Call WriteCubeRow(ws, 2, "CUBEVALUE", _
        "=CUBEVALUE(""PowerPivot Data"",""[Measures].[銷售金額]"")", _
        "從 OLAP Cube 或 Power Pivot 取得量值")
    Call WriteCubeRow(ws, 3, "CUBEMEMBER", _
        "=CUBEMEMBER(""PowerPivot Data"",""[產品].[類別].[電子]"")", _
        "從 Cube 取得成員或集合")
    Call WriteCubeRow(ws, 4, "CUBESET", _
        "=CUBESET(""PowerPivot Data"",""｛[產品].[類別].[電子],[產品].[類別].[服飾]｝"")", _
        "定義一組成員或 Tuple 的計算集合")
    Call WriteCubeRow(ws, 5, "CUBESETCOUNT", _
        "=CUBESETCOUNT(B4)", _
        "傳回集合中的項目數")
    Call WriteCubeRow(ws, 6, "CUBERANKEDMEMBER", _
        "=CUBERANKEDMEMBER(""PowerPivot Data"",B4,1)", _
        "傳回集合中排名第 N 的成員")
    Call WriteCubeRow(ws, 7, "CUBEMEMBERPROPERTY", _
        "=CUBEMEMBERPROPERTY(""PowerPivot Data"",B3,""[產品].[類別].[描述]"")", _
        "傳回 Cube 成員的屬性值")
    Call WriteCubeRow(ws, 8, "CUBEKPIMEMBER", _
        "=CUBEKPIMEMBER(""PowerPivot Data"",""[銷售 KPI]"",1)", _
        "取得 KPI 指標名稱、值或狀態")

    ws.Columns("A:C").AutoFit

    MsgBox "Cube 函數範例工作表已建立完成！" & Chr(10) & _
           "注意：公式需連線至 Power Pivot 或 SSAS 才能取得實際值。", _
           vbInformation, "完成"
    Exit Sub

ErrorHandler:
    MsgBox "建立 Cube 函數範例時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 寫入一列 Cube 函數說明
Private Sub WriteCubeRow(ByVal ws As Worksheet, ByVal rowIdx As Long, _
                          ByVal funcName As String, ByVal formula As String, _
                          ByVal desc As String)
    ws.Cells(rowIdx, 1).Value = funcName
    ws.Cells(rowIdx, 2).Value = "'" & formula
    ws.Cells(rowIdx, 3).Value = desc
End Sub

' 取得或建立工作表
Private Function GetOrCreateCubeSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateCubeSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateCubeSheet Is Nothing Then
        Set GetOrCreateCubeSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateCubeSheet.Name = sheetName
    End If
End Function
