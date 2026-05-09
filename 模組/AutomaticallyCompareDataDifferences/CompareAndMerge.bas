Attribute VB_Name = "CompareAndMerge"
Option Explicit
'*************************************************************************************
'模組名稱: CompareAndMerge
'功能說明: 比對新舊兩張工作表資料，將新版有異動的欄位更新至主表，
'          並標記異動儲存格便於覆核
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestCompareAndMerge()
    Call CreateMergeData
    Call CompareAndMerge("舊版資料", "新版資料", "合併後主表", 1, 1)
End Sub

' 建立合併測試範例資料
Private Sub CreateMergeData()
    Dim wsOld As Worksheet
    Dim wsNew As Worksheet

    Set wsOld = GetOrCreateSheetCAM("舊版資料")
    Set wsNew = GetOrCreateSheetCAM("新版資料")

    wsOld.Range("A1").Value = "產品代碼"
    wsOld.Range("B1").Value = "產品名稱"
    wsOld.Range("C1").Value = "單價"
    wsOld.Range("D1").Value = "庫存量"
    wsOld.Range("A2").Value = "P001" : wsOld.Range("B2").Value = "筆記型電腦" : wsOld.Range("C2").Value = 35000 : wsOld.Range("D2").Value = 50
    wsOld.Range("A3").Value = "P002" : wsOld.Range("B3").Value = "無線滑鼠"   : wsOld.Range("C3").Value = 800   : wsOld.Range("D3").Value = 200
    wsOld.Range("A4").Value = "P003" : wsOld.Range("B4").Value = "機械鍵盤"   : wsOld.Range("C4").Value = 2500  : wsOld.Range("D4").Value = 80
    wsOld.Columns("A:D").AutoFit

    wsNew.Range("A1").Value = "產品代碼"
    wsNew.Range("B1").Value = "產品名稱"
    wsNew.Range("C1").Value = "單價"
    wsNew.Range("D1").Value = "庫存量"
    wsNew.Range("A2").Value = "P001" : wsNew.Range("B2").Value = "筆記型電腦" : wsNew.Range("C2").Value = 38000 : wsNew.Range("D2").Value = 45
    wsNew.Range("A3").Value = "P002" : wsNew.Range("B3").Value = "無線滑鼠"   : wsNew.Range("C3").Value = 800   : wsNew.Range("D3").Value = 180
    wsNew.Range("A4").Value = "P003" : wsNew.Range("B4").Value = "機械鍵盤"   : wsNew.Range("C4").Value = 2800  : wsNew.Range("D4").Value = 80
    wsNew.Range("A5").Value = "P004" : wsNew.Range("B5").Value = "USB集線器"  : wsNew.Range("C5").Value = 450   : wsNew.Range("D5").Value = 300
    wsNew.Columns("A:D").AutoFit
End Sub

' 依鍵值欄合併新版資料至主表，並標記變更欄位
Public Sub CompareAndMerge(ByVal oldSheet As String, ByVal newSheet As String, _
                            ByVal masterSheet As String, ByVal keyColNum As Long, _
                            ByVal headerRows As Long)
    Dim wsOld        As Worksheet
    Dim wsNew        As Worksheet
    Dim wsMaster     As Worksheet
    Dim oldLastRow   As Long
    Dim newLastRow   As Long
    Dim lastCol      As Long
    Dim masterLastRow As Long
    Dim i            As Long
    Dim j            As Long
    Dim c            As Long
    Dim oldKey       As String
    Dim newKey       As String
    Dim foundMatch   As Boolean
    Dim updateCount  As Long
    Dim addCount     As Long

    On Error GoTo ErrHandler

    Set wsOld = ThisWorkbook.Worksheets(oldSheet)
    Set wsNew = ThisWorkbook.Worksheets(newSheet)
    Set wsMaster = GetOrCreateSheetCAM(masterSheet)

    ' 先複製舊版資料至主表
    wsOld.UsedRange.Copy wsMaster.Range("A1")

    oldLastRow = wsOld.Cells(wsOld.Rows.Count, keyColNum).End(xlUp).Row
    newLastRow = wsNew.Cells(wsNew.Rows.Count, keyColNum).End(xlUp).Row
    lastCol = wsOld.Cells(1, wsOld.Columns.Count).End(xlToLeft).Column

    updateCount = 0
    addCount = 0

    For i = headerRows + 1 To newLastRow
        newKey = CStr(wsNew.Cells(i, keyColNum).Value)
        If newKey <> "" Then
            foundMatch = False
            For j = headerRows + 1 To oldLastRow
                oldKey = CStr(wsOld.Cells(j, keyColNum).Value)
                If newKey = oldKey Then
                    foundMatch = True
                    ' 更新主表中有差異的欄位並加入批注
                    For c = 1 To lastCol
                        If CStr(wsNew.Cells(i, c).Value) <> CStr(wsOld.Cells(j, c).Value) Then
                            wsMaster.Cells(j, c).Value = wsNew.Cells(i, c).Value
                            wsMaster.Cells(j, c).Interior.Color = RGB(146, 208, 80)
                            On Error Resume Next
                            wsMaster.Cells(j, c).AddComment "原值: " & CStr(wsOld.Cells(j, c).Value)
                            On Error GoTo ErrHandler
                            updateCount = updateCount + 1
                        End If
                    Next c
                    Exit For
                End If
            Next j
            If Not foundMatch Then
                ' 新增列加入主表底部
                masterLastRow = wsMaster.Cells(wsMaster.Rows.Count, keyColNum).End(xlUp).Row + 1
                For c = 1 To lastCol
                    wsMaster.Cells(masterLastRow, c).Value = wsNew.Cells(i, c).Value
                Next c
                wsMaster.Rows(masterLastRow).Interior.Color = RGB(198, 239, 206)
                addCount = addCount + 1
            End If
        End If
    Next i

    wsMaster.Columns("A:D").AutoFit
    wsMaster.Activate
    MsgBox "合併完成！" & vbCrLf & _
           "更新欄位: " & updateCount & " 格" & vbCrLf & _
           "新增記錄: " & addCount & " 筆", vbInformation, "合併結果"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetCAM(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetCAM = ws
End Function
