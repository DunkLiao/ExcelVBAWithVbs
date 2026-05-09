Attribute VB_Name = "CompareByKeyColumn"
Option Explicit
'*************************************************************************************
'模組名稱: CompareByKeyColumn
'功能說明: 依指定鍵值欄位比對兩張工作表，輸出新增、刪除、修改三類差異
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestCompareByKeyColumn()
    Call CreateKeyCompareData
    Call CompareByKeyColumn("來源清單", "目標清單", 1, "鍵值差異報告", 1)
End Sub

' 建立鍵值比對範例資料
Private Sub CreateKeyCompareData()
    Dim wsA As Worksheet
    Dim wsB As Worksheet

    Set wsA = GetOrCreateSheetCBK("來源清單")
    Set wsB = GetOrCreateSheetCBK("目標清單")

    wsA.Range("A1").Value = "訂單編號"
    wsA.Range("B1").Value = "客戶"
    wsA.Range("C1").Value = "金額"
    wsA.Range("A2").Value = "ORD001" : wsA.Range("B2").Value = "台積電" : wsA.Range("C2").Value = 50000
    wsA.Range("A3").Value = "ORD002" : wsA.Range("B3").Value = "聯發科" : wsA.Range("C3").Value = 32000
    wsA.Range("A4").Value = "ORD003" : wsA.Range("B4").Value = "鴻海"   : wsA.Range("C4").Value = 18000
    wsA.Columns("A:C").AutoFit

    wsB.Range("A1").Value = "訂單編號"
    wsB.Range("B1").Value = "客戶"
    wsB.Range("C1").Value = "金額"
    wsB.Range("A2").Value = "ORD001" : wsB.Range("B2").Value = "台積電" : wsB.Range("C2").Value = 55000
    wsB.Range("A3").Value = "ORD002" : wsB.Range("B3").Value = "聯發科" : wsB.Range("C3").Value = 32000
    wsB.Range("A4").Value = "ORD004" : wsB.Range("B4").Value = "華碩"   : wsB.Range("C4").Value = 41000
    wsB.Columns("A:C").AutoFit
End Sub

' 依鍵值欄比對，輸出新增/刪除/修改清單
Public Sub CompareByKeyColumn(ByVal srcSheet As String, ByVal tgtSheet As String, _
                               ByVal keyColNum As Long, ByVal reportSheet As String, _
                               ByVal headerRows As Long)
    Dim wsS         As Worksheet
    Dim wsT         As Worksheet
    Dim wsR         As Worksheet
    Dim srcLastRow  As Long
    Dim tgtLastRow  As Long
    Dim srcLastCol  As Long
    Dim rptRow      As Long
    Dim i           As Long
    Dim j           As Long
    Dim c           As Long
    Dim srcKey      As String
    Dim tgtKey      As String
    Dim foundInTarget As Boolean
    Dim foundInSource As Boolean
    Dim addedCount  As Long
    Dim deletedCount As Long
    Dim modifiedCount As Long

    On Error GoTo ErrHandler

    Set wsS = ThisWorkbook.Worksheets(srcSheet)
    Set wsT = ThisWorkbook.Worksheets(tgtSheet)
    Set wsR = GetOrCreateSheetCBK(reportSheet)

    srcLastRow = wsS.Cells(wsS.Rows.Count, keyColNum).End(xlUp).Row
    tgtLastRow = wsT.Cells(wsT.Rows.Count, keyColNum).End(xlUp).Row
    srcLastCol = wsS.Cells(1, wsS.Columns.Count).End(xlToLeft).Column

    wsR.Range("A1").Value = "差異類型"
    wsR.Range("B1").Value = "鍵值"
    wsR.Range("C1").Value = "欄位"
    wsR.Range("D1").Value = "來源值"
    wsR.Range("E1").Value = "目標值"
    With wsR.Range("A1:E1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    rptRow = 2
    addedCount = 0
    deletedCount = 0
    modifiedCount = 0

    ' 掃描來源，找出修改與刪除
    For i = headerRows + 1 To srcLastRow
        srcKey = CStr(wsS.Cells(i, keyColNum).Value)
        If srcKey <> "" Then
            foundInTarget = False
            For j = headerRows + 1 To tgtLastRow
                tgtKey = CStr(wsT.Cells(j, keyColNum).Value)
                If srcKey = tgtKey Then
                    foundInTarget = True
                    For c = 1 To srcLastCol
                        If CStr(wsS.Cells(i, c).Value) <> CStr(wsT.Cells(j, c).Value) Then
                            wsR.Cells(rptRow, 1).Value = "修改"
                            wsR.Cells(rptRow, 2).Value = srcKey
                            wsR.Cells(rptRow, 3).Value = CStr(wsS.Cells(1, c).Value)
                            wsR.Cells(rptRow, 4).Value = CStr(wsS.Cells(i, c).Value)
                            wsR.Cells(rptRow, 5).Value = CStr(wsT.Cells(j, c).Value)
                            wsR.Cells(rptRow, 1).Interior.Color = RGB(255, 255, 153)
                            rptRow = rptRow + 1
                            modifiedCount = modifiedCount + 1
                        End If
                    Next c
                    Exit For
                End If
            Next j
            If Not foundInTarget Then
                wsR.Cells(rptRow, 1).Value = "刪除"
                wsR.Cells(rptRow, 2).Value = srcKey
                wsR.Cells(rptRow, 3).Value = "(不存在於目標)"
                wsR.Cells(rptRow, 1).Interior.Color = RGB(255, 199, 206)
                rptRow = rptRow + 1
                deletedCount = deletedCount + 1
            End If
        End If
    Next i

    ' 掃描目標，找出新增
    For j = headerRows + 1 To tgtLastRow
        tgtKey = CStr(wsT.Cells(j, keyColNum).Value)
        If tgtKey <> "" Then
            foundInSource = False
            For i = headerRows + 1 To srcLastRow
                If tgtKey = CStr(wsS.Cells(i, keyColNum).Value) Then
                    foundInSource = True
                    Exit For
                End If
            Next i
            If Not foundInSource Then
                wsR.Cells(rptRow, 1).Value = "新增"
                wsR.Cells(rptRow, 2).Value = tgtKey
                wsR.Cells(rptRow, 3).Value = "(僅存在於目標)"
                wsR.Cells(rptRow, 1).Interior.Color = RGB(198, 239, 206)
                rptRow = rptRow + 1
                addedCount = addedCount + 1
            End If
        End If
    Next j

    wsR.Columns("A:E").AutoFit
    wsR.Activate
    MsgBox "鍵值比對完成！" & vbCrLf & _
           "新增: " & addedCount & " 筆" & vbCrLf & _
           "刪除: " & deletedCount & " 筆" & vbCrLf & _
           "修改: " & modifiedCount & " 筆", vbInformation, "比對結果"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetCBK(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetCBK = ws
End Function
