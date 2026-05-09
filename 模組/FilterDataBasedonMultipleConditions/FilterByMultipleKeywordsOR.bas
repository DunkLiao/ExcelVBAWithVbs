Attribute VB_Name = "FilterByMultipleKeywordsOR"
Option Explicit

'************************************************************************************
' 模組名稱: FilterByMultipleKeywordsOR
' 功能說明: 使用 AdvancedFilter OR 邏輯，篩選某欄符合多個關鍵字之任一的資料
'           每個關鍵字條件分列在條件區的不同列（OR 關係）
'
' 作者版權: Dunk
' 現任設計: Dunk
' 最後修改: 2026/5/9
'************************************************************************************

' 入口：篩選「城市」欄中包含台北、台中或高雄的客戶
Public Sub FilterByMultipleKeywordsORExample()
    On Error GoTo ErrHandler

    Dim wsData     As Worksheet
    Dim wsCriteria As Worksheet
    Dim wsResult   As Worksheet

    Set wsData = GetOrCreateWsKw(ThisWorkbook, "多關鍵字來源")
    Call FillCityData(wsData)

    ' 建立條件區（OR：各關鍵字分行）
    Set wsCriteria = GetOrCreateWsKw(ThisWorkbook, "多關鍵字條件")
    wsCriteria.Range("A1").Value = "城市"
    wsCriteria.Range("A2").Value = "台北*"
    wsCriteria.Range("A3").Value = "台中*"
    wsCriteria.Range("A4").Value = "高雄*"

    Set wsResult = GetOrCreateWsKw(ThisWorkbook, "多關鍵字結果")

    wsData.Range("A1").CurrentRegion.AdvancedFilter _
        Action:=xlFilterCopy, _
        CriteriaRange:=wsCriteria.Range("A1:A4"), _
        CopyToRange:=wsResult.Range("A1"), _
        Unique:=False

    wsResult.Columns.AutoFit
    wsResult.Activate
    MsgBox "已篩選台北/台中/高雄客戶，結果在「多關鍵字結果」工作表。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 填入客戶城市測試資料
Private Sub FillCityData(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("客戶名稱", "城市", "消費金額")
    ws.Range("A2:C2").Value = Array("甲公司", "台北市信義區", 85000)
    ws.Range("A3:C3").Value = Array("乙公司", "新竹市東區", 42000)
    ws.Range("A4:C4").Value = Array("丙公司", "台中市西屯區", 67000)
    ws.Range("A5:C5").Value = Array("丁公司", "高雄市前鎮區", 91000)
    ws.Range("A6:C6").Value = Array("戊公司", "桃園市中壢區", 35000)
    ws.Range("A7:C7").Value = Array("己公司", "台北市大安區", 120000)
    ws.Range("A8:C8").Value = Array("庚公司", "台南市永康區", 58000)
    ws.Range("A9:C9").Value = Array("辛公司", "高雄市鹽埕區", 47000)
    ws.Range("C2:C9").NumberFormat = "#,##0"
    ws.Columns("A:C").AutoFit
End Sub

' 取得或建立工作表並清空
Private Function GetOrCreateWsKw(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateWsKw = ws
End Function