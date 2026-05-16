Attribute VB_Name = "FilterByMultipleSheets"
Option Explicit
'*************************************************************************************
'模組名稱: FilterByMultipleSheets
'功能說明: 從多個工作表各取一欄作為條件清單，對主資料工作表進行
'          多工作表條件交集篩選，結果輸出至新工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/16
'
'*************************************************************************************

' 主程式：依多個工作表的條件清單篩選主資料
Sub FilterDataByMultipleSheetConditions()
    Dim wbThis       As Workbook
    Dim wsMain       As Worksheet
    Dim wsResult     As Worksheet
    Dim mainName     As String
    Dim resultName   As String
    Dim mainLastRow  As Long
    Dim mainLastCol  As Long
    Dim r            As Long
    Dim c            As Long
    Dim colIdx       As Long
    Dim destRow      As Long
    Dim matched      As Boolean

    Set wbThis = ThisWorkbook

    ' 取得主資料工作表名稱
    mainName = InputBox("請輸入主資料工作表名稱：", "設定", "主資料")
    If mainName = "" Then Exit Sub

    On Error Resume Next
    Set wsMain = wbThis.Worksheets(mainName)
    On Error GoTo 0

    If wsMain Is Nothing Then
        MsgBox "找不到工作表「" & mainName & "」。", vbExclamation
        Exit Sub
    End If

    mainLastRow = wsMain.Cells(wsMain.Rows.Count, 1).End(xlUp).Row
    mainLastCol = wsMain.Cells(1, wsMain.Columns.Count).End(xlToLeft).Column

    If mainLastRow < 2 Then
        MsgBox "主資料工作表沒有資料列。", vbExclamation
        Exit Sub
    End If

    ' 收集每個條件工作表的設定（工作表名稱 + 條件欄位 + 主表欄位）
    Dim condSheet(1 To 5) As String
    Dim condColMain(1 To 5) As Long  ' 主表的對應欄索引
    Dim condCount    As Integer
    condCount = 0

    Dim addMore As Integer
    addMore = vbYes

    Do While addMore = vbYes And condCount < 5
        Dim sName As String
        sName = InputBox("請輸入條件工作表名稱（清單位於 A 欄，含標題）：", _
                         "條件工作表 " & (condCount + 1))
        If sName = "" Then Exit Do

        Dim wsCheck As Worksheet
        On Error Resume Next
        Set wsCheck = wbThis.Worksheets(sName)
        On Error GoTo 0

        If wsCheck Is Nothing Then
            MsgBox "找不到工作表「" & sName & "」，請重新輸入。", vbExclamation
        Else
            Dim matchColStr As String
            matchColStr = InputBox("條件工作表「" & sName & "」的 A 欄要對應主資料的哪一欄？" & _
                                   "（請輸入欄位代號，例如 B）：", "欄位對應", "B")
            If matchColStr = "" Then
                MsgBox "已取消此條件。", vbInformation
            Else
                condCount = condCount + 1
                condSheet(condCount) = sName
                condColMain(condCount) = wsMain.Range(matchColStr & "1").Column
            End If
        End If

        If condCount < 5 Then
            addMore = MsgBox("是否繼續新增條件工作表？", vbYesNo + vbQuestion, "繼續新增")
        Else
            addMore = vbNo
        End If
    Loop

    If condCount = 0 Then
        MsgBox "未設定任何條件，已結束。", vbInformation
        Exit Sub
    End If

    ' 建立各條件工作表的條件集合（Dictionary）
    Dim dictConds(1 To 5) As Object
    Dim i As Integer
    For i = 1 To condCount
        Set dictConds(i) = CreateObject("Scripting.Dictionary")
        Dim wsC As Worksheet
        Set wsC = wbThis.Worksheets(condSheet(i))
        Dim cLastRow As Long
        cLastRow = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
        Dim cr As Long
        For cr = 2 To cLastRow
            Dim condVal As String
            condVal = Trim(CStr(wsC.Cells(cr, 1).Value))
            If condVal <> "" And Not dictConds(i).Exists(condVal) Then
                dictConds(i).Add condVal, condVal
            End If
        Next cr
    Next i

    ' 建立結果工作表
    resultName = "篩選結果_多表"
    Dim wsOld As Worksheet
    On Error Resume Next
    Set wsOld = wbThis.Worksheets(resultName)
    On Error GoTo 0
    If Not wsOld Is Nothing Then
        Application.DisplayAlerts = False
        wsOld.Delete
        Application.DisplayAlerts = True
    End If

    Set wsResult = wbThis.Worksheets.Add
    wsResult.Name = resultName

    ' 複製標題列
    wsMain.Range(wsMain.Cells(1, 1), wsMain.Cells(1, mainLastCol)).Copy _
          wsResult.Range("A1")
    destRow = 2

    Application.ScreenUpdating = False

    ' 逐列比對條件
    For r = 2 To mainLastRow
        matched = True
        For i = 1 To condCount
            Dim cellVal As String
            cellVal = Trim(CStr(wsMain.Cells(r, condColMain(i)).Value))
            If Not dictConds(i).Exists(cellVal) Then
                matched = False
                Exit For
            End If
        Next i

        If matched Then
            wsMain.Range(wsMain.Cells(r, 1), wsMain.Cells(r, mainLastCol)).Copy _
                  wsResult.Range("A" & destRow)
            destRow = destRow + 1
        End If
    Next r

    Application.ScreenUpdating = True
    wsResult.Columns.AutoFit

    Dim resultCount As Long
    resultCount = destRow - 2

    MsgBox "篩選完成！符合所有條件的資料共 " & resultCount & " 筆，" & _
           "已輸出至工作表「" & resultName & "」。", vbInformation, "完成"
End Sub

' 建立示範並測試功能
Sub DemoFilterByMultipleSheets()
    Dim ws As Worksheet

    ' 建立主資料工作表
    Set ws = GetOrCreateMultiSheet("主資料")
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("業務員", "地區", "產品")
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A2:C2").Value = Array("張志豪", "北區", "A產品")
    ws.Range("A3:C3").Value = Array("李佳蓉", "中區", "B產品")
    ws.Range("A4:C4").Value = Array("王大明", "南區", "A產品")
    ws.Range("A5:C5").Value = Array("張志豪", "北區", "B產品")
    ws.Range("A6:C6").Value = Array("陳雅婷", "中區", "A產品")
    ws.Range("A7:C7").Value = Array("林志偉", "北區", "A產品")
    ws.Columns("A:C").AutoFit

    ' 建立地區條件工作表
    Dim wsR As Worksheet
    Set wsR = GetOrCreateMultiSheet("地區條件")
    wsR.Cells.Clear
    wsR.Range("A1").Value = "地區"
    wsR.Range("A2").Value = "北區"
    wsR.Range("A3").Value = "中區"

    ' 建立產品條件工作表
    Dim wsP As Worksheet
    Set wsP = GetOrCreateMultiSheet("產品條件")
    wsP.Cells.Clear
    wsP.Range("A1").Value = "產品"
    wsP.Range("A2").Value = "A產品"

    MsgBox "示範資料已建立！" & vbCrLf & _
           "主資料：「主資料」工作表" & vbCrLf & _
           "條件1：「地區條件」（北區、中區）" & vbCrLf & _
           "條件2：「產品條件」（A產品）" & vbCrLf & vbCrLf & _
           "請執行 FilterDataByMultipleSheetConditions 並依提示輸入工作表名稱。", _
           vbInformation, "示範說明"
End Sub

Private Function GetOrCreateMultiSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateMultiSheet = ws
End Function
