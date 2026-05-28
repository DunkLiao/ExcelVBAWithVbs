Attribute VB_Name = "FilterByExactMatchList"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByExactMatchList
'功能說明: 依精確比對清單篩選工作表資料，只保留指定清單中完全相符的列，
'          結果複製至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/28
'
'*************************************************************************************

Sub TestFilterByExactMatchList()
    Call CreateExactMatchDemo(ThisWorkbook)
    Dim srcWs  As Worksheet
    Dim listWs As Worksheet
    Set srcWs  = ThisWorkbook.Worksheets("精確比對來源")
    Set listWs = ThisWorkbook.Worksheets("比對清單")
    Call FilterByExactMatchList(srcWs, listWs, 2, 1, "精確比對結果")
End Sub

Private Sub CreateExactMatchDemo(ByVal wb As Workbook)
    Dim srcWs  As Worksheet
    Dim listWs As Worksheet
    On Error Resume Next
    Set srcWs  = wb.Worksheets("精確比對來源")
    Set listWs = wb.Worksheets("比對清單")
    On Error GoTo 0
    If srcWs Is Nothing Then
        Set srcWs = wb.Worksheets.Add
        srcWs.Name = "精確比對來源"
    End If
    srcWs.Cells.Clear
    If listWs Is Nothing Then
        Set listWs = wb.Worksheets.Add(After:=srcWs)
        listWs.Name = "比對清單"
    End If
    listWs.Cells.Clear

    srcWs.Range("A1:D1").Value = Array("姓名", "城市", "部門", "業績")
    srcWs.Range("A2:D2").Value = Array("張小明", "台北", "業務部", 320000)
    srcWs.Range("A3:D3").Value = Array("李美玲", "台中", "財務部", 280000)
    srcWs.Range("A4:D4").Value = Array("王大偉", "高雄", "工程部", 410000)
    srcWs.Range("A5:D5").Value = Array("陳佳芳", "台北", "人資部", 190000)
    srcWs.Range("A6:D6").Value = Array("林建志", "新竹", "業務部", 355000)
    srcWs.Range("A7:D7").Value = Array("吳雅雯", "台中", "工程部", 475000)
    srcWs.Range("A8:D8").Value = Array("黃志偉", "台南", "財務部", 230000)
    srcWs.Range("A9:D9").Value = Array("許淑芬", "台北", "業務部", 390000)
    srcWs.Range("A10:D10").Value = Array("鄭宏達", "桃園", "工程部", 510000)
    srcWs.Range("A11:D11").Value = Array("謝明輝", "高雄", "人資部", 175000)
    srcWs.Range("A12:D12").Value = Array("蔡欣怡", "新竹", "財務部", 295000)
    srcWs.Range("A13:D13").Value = Array("周建國", "台南", "業務部", 340000)
    srcWs.Columns("A:D").AutoFit

    listWs.Range("A1").Value = "要篩選的城市"
    listWs.Range("A2").Value = "台北"
    listWs.Range("A3").Value = "高雄"
    listWs.Range("A4").Value = "新竹"
    listWs.Columns("A").AutoFit
End Sub

Sub FilterByExactMatchList(ByVal srcWs As Worksheet, _
                            ByVal listWs As Worksheet, _
                            ByVal filterCol As Integer, _
                            ByVal listCol As Integer, _
                            ByVal resultName As String)
    Dim resultWs    As Worksheet
    Dim matchDict   As Object
    Dim lastSrcRow  As Long
    Dim lastLstRow  As Long
    Dim i           As Long
    Dim matchVal    As String
    Dim resultRow   As Long

    Set matchDict = CreateObject("Scripting.Dictionary")
    matchDict.CompareMode = 1

    lastLstRow = listWs.Cells(listWs.Rows.Count, listCol).End(xlUp).Row
    For i = 2 To lastLstRow
        matchVal = CStr(listWs.Cells(i, listCol).Value)
        If Len(matchVal) > 0 And Not matchDict.Exists(matchVal) Then
            matchDict.Add matchVal, True
        End If
    Next i

    If matchDict.Count = 0 Then
        MsgBox "比對清單為空，無法執行篩選。", vbExclamation, "錯誤"
        Exit Sub
    End If

    On Error Resume Next
    Set resultWs = srcWs.Parent.Worksheets(resultName)
    On Error GoTo 0
    If resultWs Is Nothing Then
        Set resultWs = srcWs.Parent.Worksheets.Add( _
            After:=srcWs.Parent.Worksheets(srcWs.Parent.Worksheets.Count))
        resultWs.Name = resultName
    End If
    resultWs.Cells.Clear

    lastSrcRow = srcWs.Cells(srcWs.Rows.Count, filterCol).End(xlUp).Row
    srcWs.Rows(1).Copy
    resultWs.Rows(1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    resultRow = 2

    Application.ScreenUpdating = False
    For i = 2 To lastSrcRow
        matchVal = CStr(srcWs.Cells(i, filterCol).Value)
        If matchDict.Exists(matchVal) Then
            srcWs.Rows(i).Copy
            resultWs.Rows(resultRow).PasteSpecial Paste:=xlPasteValues
            Application.CutCopyMode = False
            resultRow = resultRow + 1
        End If
    Next i

    resultWs.Columns.AutoFit
    resultWs.Activate
    Application.ScreenUpdating = True
    MsgBox "精確比對篩選完成！" & vbCrLf & _
           "比對清單：" & matchDict.Count & " 個值" & vbCrLf & _
           "篩選結果：" & (resultRow - 2) & " 筆資料。", _
           vbInformation, "完成"
End Sub
