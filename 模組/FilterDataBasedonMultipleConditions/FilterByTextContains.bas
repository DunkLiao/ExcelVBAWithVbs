Attribute VB_Name = "FilterByTextContains"
Option Explicit

'************************************************************************************
' 模組名稱: FilterByTextContains
' 功能說明: 使用 AutoFilter 萬用字元篩選包含特定文字的資料列
'           示範 * 萬用字元搭配多欄條件篩選
'
' 作者版權: Dunk
' 現任設計: Dunk
' 最後修改: 2026/5/9
'************************************************************************************

' 主要入口：篩選品名含「電」且供應商含「科技」的資料
Public Sub FilterByTextContainsExample()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = GetOrCreateWs(ThisWorkbook, "文字篩選範例")
    Call FillProductData(ws)
    Call ApplyTextContainsFilter(ws, "電", "科技")

    MsgBox "篩選完成！已保留品名含「電」且供應商含「科技」的列。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

' 套用文字包含篩選
Private Sub ApplyTextContainsFilter(ByVal ws As Worksheet, _
                                     ByVal productKeyword As String, _
                                     ByVal supplierKeyword As String)
    Dim rng As Range
    Set rng = ws.Range("A1").CurrentRegion

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' Field:=1 品名含 productKeyword，Field:=3 供應商含 supplierKeyword
    rng.AutoFilter Field:=1, Criteria1:="*" & productKeyword & "*"
    rng.AutoFilter Field:=3, Criteria1:="*" & supplierKeyword & "*"
    ws.Columns("A:D").AutoFit
End Sub

' 填入商品測試資料
Private Sub FillProductData(ByVal ws As Worksheet)
    ws.Range("A1:D1").Value = Array("品名", "單價", "供應商", "庫存")
    ws.Range("A2:D2").Value = Array("液晶電視 55吋", 18000, "新光科技", 30)
    ws.Range("A3:D3").Value = Array("冷氣機 1.5噸", 25000, "東元電器", 20)
    ws.Range("A4:D4").Value = Array("電磁爐 IH", 3200, "品佳科技", 80)
    ws.Range("A5:D5").Value = Array("掃地機器人", 8900, "小米科技", 50)
    ws.Range("A6:D6").Value = Array("電風扇 DC", 2200, "大同電器", 120)
    ws.Range("A7:D7").Value = Array("投影機 4K", 35000, "優派科技", 10)
    ws.Range("A8:D8").Value = Array("洗衣機 滾筒", 22000, "日立電器", 15)
    ws.Columns("A:D").AutoFit
End Sub

' 取得或建立工作表並清空
Private Function GetOrCreateWs(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = shName
    End If
    ws.Cells.Clear
    Set GetOrCreateWs = ws
End Function