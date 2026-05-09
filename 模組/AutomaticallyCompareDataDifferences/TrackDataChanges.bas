Attribute VB_Name = "TrackDataChanges"
Option Explicit
'*************************************************************************************
'模組名稱: TrackDataChanges
'功能說明: 比對兩版資料，將差異記錄至變更歷程工作表，
'          包含變更時間、變更前後值、變更類型
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestTrackDataChanges()
    Call CreateTrackData
    Call TrackDataChanges("v1.0資料", "v2.0資料", "變更歷程", 1, 1)
End Sub

' 建立版本追蹤範例資料
Private Sub CreateTrackData()
    Dim wsV1 As Worksheet
    Dim wsV2 As Worksheet

    Set wsV1 = GetOrCreateSheetTDC("v1.0資料")
    Set wsV2 = GetOrCreateSheetTDC("v2.0資料")

    wsV1.Range("A1").Value = "客戶編號"
    wsV1.Range("B1").Value = "公司名稱"
    wsV1.Range("C1").Value = "聯絡人"
    wsV1.Range("D1").Value = "信用等級"
    wsV1.Range("A2").Value = "C001" : wsV1.Range("B2").Value = "宏碁股份有限公司"   : wsV1.Range("C2").Value = "陳小明" : wsV1.Range("D2").Value = "A"
    wsV1.Range("A3").Value = "C002" : wsV1.Range("B3").Value = "微星科技股份有限公司" : wsV1.Range("C3").Value = "李大華" : wsV1.Range("D3").Value = "B"
    wsV1.Range("A4").Value = "C003" : wsV1.Range("B4").Value = "技嘉科技股份有限公司" : wsV1.Range("C4").Value = "王雅惠" : wsV1.Range("D4").Value = "A"
    wsV1.Columns("A:D").AutoFit

    wsV2.Range("A1").Value = "客戶編號"
    wsV2.Range("B1").Value = "公司名稱"
    wsV2.Range("C1").Value = "聯絡人"
    wsV2.Range("D1").Value = "信用等級"
    wsV2.Range("A2").Value = "C001" : wsV2.Range("B2").Value = "宏碁股份有限公司"   : wsV2.Range("C2").Value = "陳小明" : wsV2.Range("D2").Value = "A+"
    wsV2.Range("A3").Value = "C002" : wsV2.Range("B3").Value = "微星科技股份有限公司" : wsV2.Range("C3").Value = "林淑芬" : wsV2.Range("D3").Value = "B"
    wsV2.Range("A4").Value = "C003" : wsV2.Range("B4").Value = "技嘉科技股份有限公司" : wsV2.Range("C4").Value = "王雅惠" : wsV2.Range("D4").Value = "A+"
    wsV2.Range("A5").Value = "C004" : wsV2.Range("B5").Value = "仁寶電腦工業股份有限公司" : wsV2.Range("C5").Value = "黃志偉" : wsV2.Range("D5").Value = "B+"
    wsV2.Columns("A:D").AutoFit
End Sub

' 追蹤並記錄兩版資料間的變更歷程
Public Sub TrackDataChanges(ByVal oldVersion As String, ByVal newVersion As String, _
                             ByVal logSheet As String, ByVal keyColNum As Long, _
                             ByVal headerRows As Long)
    Dim wsOld      As Worksheet
    Dim wsNew      As Worksheet
    Dim wsLog      As Worksheet
    Dim oldLastRow As Long
    Dim newLastRow As Long
    Dim lastCol    As Long
    Dim i          As Long
    Dim j          As Long
    Dim c          As Long
    Dim oldKey     As String
    Dim newKey     As String
    Dim foundMatch As Boolean
    Dim logRow     As Long
    Dim changeTime As String

    On Error GoTo ErrHandler

    Set wsOld = ThisWorkbook.Worksheets(oldVersion)
    Set wsNew = ThisWorkbook.Worksheets(newVersion)
    Set wsLog = GetOrCreateSheetTDC(logSheet)

    changeTime = Format(Now, "yyyy/mm/dd hh:mm:ss")

    oldLastRow = wsOld.Cells(wsOld.Rows.Count, keyColNum).End(xlUp).Row
    newLastRow = wsNew.Cells(wsNew.Rows.Count, keyColNum).End(xlUp).Row
    lastCol = wsOld.Cells(1, wsOld.Columns.Count).End(xlToLeft).Column

    ' 設定歷程標頭
    wsLog.Range("A1").Value = "記錄時間"
    wsLog.Range("B1").Value = "變更類型"
    wsLog.Range("C1").Value = "鍵值"
    wsLog.Range("D1").Value = "欄位名稱"
    wsLog.Range("E1").Value = "舊版值"
    wsLog.Range("F1").Value = "新版值"
    wsLog.Range("G1").Value = "版本對照"
    With wsLog.Range("A1:G1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 73, 125)
        .Font.Color = RGB(255, 255, 255)
    End With

    logRow = 2

    ' 比對舊版 vs 新版（修改 / 刪除）
    For i = headerRows + 1 To oldLastRow
        oldKey = CStr(wsOld.Cells(i, keyColNum).Value)
        If oldKey <> "" Then
            foundMatch = False
            For j = headerRows + 1 To newLastRow
                newKey = CStr(wsNew.Cells(j, keyColNum).Value)
                If oldKey = newKey Then
                    foundMatch = True
                    For c = 1 To lastCol
                        If CStr(wsOld.Cells(i, c).Value) <> CStr(wsNew.Cells(j, c).Value) Then
                            wsLog.Cells(logRow, 1).Value = changeTime
                            wsLog.Cells(logRow, 2).Value = "修改"
                            wsLog.Cells(logRow, 3).Value = oldKey
                            wsLog.Cells(logRow, 4).Value = CStr(wsOld.Cells(1, c).Value)
                            wsLog.Cells(logRow, 5).Value = CStr(wsOld.Cells(i, c).Value)
                            wsLog.Cells(logRow, 6).Value = CStr(wsNew.Cells(j, c).Value)
                            wsLog.Cells(logRow, 7).Value = oldVersion & " -> " & newVersion
                            wsLog.Cells(logRow, 2).Interior.Color = RGB(255, 255, 153)
                            logRow = logRow + 1
                        End If
                    Next c
                    Exit For
                End If
            Next j
            If Not foundMatch Then
                wsLog.Cells(logRow, 1).Value = changeTime
                wsLog.Cells(logRow, 2).Value = "刪除"
                wsLog.Cells(logRow, 3).Value = oldKey
                wsLog.Cells(logRow, 4).Value = "(整筆記錄)"
                wsLog.Cells(logRow, 5).Value = "(存在)"
                wsLog.Cells(logRow, 6).Value = "(已刪除)"
                wsLog.Cells(logRow, 7).Value = oldVersion & " -> " & newVersion
                wsLog.Cells(logRow, 2).Interior.Color = RGB(255, 199, 206)
                logRow = logRow + 1
            End If
        End If
    Next i

    ' 找出新增記錄
    For j = headerRows + 1 To newLastRow
        newKey = CStr(wsNew.Cells(j, keyColNum).Value)
        If newKey <> "" Then
            foundMatch = False
            For i = headerRows + 1 To oldLastRow
                If newKey = CStr(wsOld.Cells(i, keyColNum).Value) Then
                    foundMatch = True
                    Exit For
                End If
            Next i
            If Not foundMatch Then
                wsLog.Cells(logRow, 1).Value = changeTime
                wsLog.Cells(logRow, 2).Value = "新增"
                wsLog.Cells(logRow, 3).Value = newKey
                wsLog.Cells(logRow, 4).Value = "(整筆記錄)"
                wsLog.Cells(logRow, 5).Value = "(不存在)"
                wsLog.Cells(logRow, 6).Value = "(已新增)"
                wsLog.Cells(logRow, 7).Value = oldVersion & " -> " & newVersion
                wsLog.Cells(logRow, 2).Interior.Color = RGB(198, 239, 206)
                logRow = logRow + 1
            End If
        End If
    Next j

    wsLog.Columns("A:G").AutoFit
    wsLog.Activate
    MsgBox "變更歷程記錄完成！共記錄 " & (logRow - 2) & " 筆變更。", vbInformation, "追蹤結果"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetTDC(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetTDC = ws
End Function
