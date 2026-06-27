Attribute VB_Name = "SplitByDataValidationList"
Option Explicit
'*************************************************************************************
'模組名稱: SplitByDataValidationList
'功能說明: 根據資料驗證清單（下拉選單值）將資料分割至不同工作表的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

' 簡化測試入口
Sub TestSplitByDataValidationList()
    Call SplitByDataValidationList
End Sub

Sub SplitByDataValidationList()
    Dim wsSource As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim category As String
    Dim catDict As Object
    Dim catKey As Variant
    Dim targetWs As Worksheet
    Dim targetRow As Long
    Dim sheetName As String
    Dim catCount As Long
    
    sheetName = "驗證清單分割"
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        Set wsSource = ThisWorkbook.Worksheets.Add
        wsSource.Name = sheetName
    End If
    
    wsSource.Cells.Clear
    Call FillValidationData(wsSource)
    
    Set catDict = CreateObject("Scripting.Dictionary")
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastRow
        category = CStr(wsSource.Cells(i, 3).Value)
        If Not catDict.Exists(category) Then
            catDict.Add category, True
        End If
    Next i
    
    catCount = catDict.Count
    
    For Each catKey In catDict.Keys
        On Error Resume Next
        ThisWorkbook.Worksheets(CStr(catKey)).Delete
        On Error GoTo 0
        
        Set targetWs = ThisWorkbook.Worksheets.Add
        targetWs.Name = CStr(catKey)
        wsSource.Rows(1).Copy targetWs.Rows(1)
        targetRow = 1
        
        For i = 2 To lastRow
            If CStr(wsSource.Cells(i, 3).Value) = CStr(catKey) Then
                targetRow = targetRow + 1
                wsSource.Rows(i).Copy targetWs.Rows(targetRow)
            End If
        Next i
        
        targetWs.Columns("A:C").AutoFit
    Next catKey
    
    wsSource.Activate
    Set catDict = Nothing
    
    MsgBox "資料已依驗證清單分類分割完成！" & vbCrLf & _
           "共分割為 " & catCount & " 個工作表。", vbInformation, "完成"
End Sub

Private Sub FillValidationData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "產品編號"
    ws.Range("B1").Value = "產品名稱"
    ws.Range("C1").Value = "產品類別"
    
    ws.Range("A2").Value = "P001"
    ws.Range("B2").Value = "筆記型電腦"
    ws.Range("C2").Value = "電子產品"
    
    ws.Range("A3").Value = "P002"
    ws.Range("B3").Value = "辦公桌"
    ws.Range("C3").Value = "家具"
    
    ws.Range("A4").Value = "P003"
    ws.Range("B4").Value = "智慧手機"
    ws.Range("C4").Value = "電子產品"
    
    ws.Range("A5").Value = "P004"
    ws.Range("B5").Value = "辦公椅"
    ws.Range("C5").Value = "家具"
    
    ws.Range("A6").Value = "P005"
    ws.Range("B6").Value = "平板電腦"
    ws.Range("C6").Value = "電子產品"
    
    ws.Range("A7").Value = "P006"
    ws.Range("B7").Value = "書櫃"
    ws.Range("C7").Value = "家具"
    
    ws.Range("A8").Value = "P007"
    ws.Range("B8").Value = "印表機"
    ws.Range("C8").Value = "辦公設備"
    
    ws.Range("A9").Value = "P008"
    ws.Range("B9").Value = "投影機"
    ws.Range("C9").Value = "辦公設備"
    
    ws.Columns("A:C").AutoFit
End Sub
