Attribute VB_Name = "CompareColumnHeaders"
Option Explicit
'*************************************************************************************
'模組名稱: CompareColumnHeaders
'功能說明: 比對兩張工作表的欄位標題結構是否一致，
'          找出缺少、多餘、順序不符的欄位
'
'著作權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/9
'
'*************************************************************************************

' 測試用入口
Sub TestCompareColumnHeaders()
    Call CreateHeaderData
    Call CompareColumnHeaders("標準格式", "上傳格式", "欄位結構差異")
End Sub

' 建立欄位標題比對範例資料
Private Sub CreateHeaderData()
    Dim wsA As Worksheet
    Dim wsB As Worksheet

    Set wsA = GetOrCreateSheetCCH("標準格式")
    Set wsB = GetOrCreateSheetCCH("上傳格式")

    wsA.Range("A1").Value = "員工編號"
    wsA.Range("B1").Value = "姓名"
    wsA.Range("C1").Value = "部門"
    wsA.Range("D1").Value = "職稱"
    wsA.Range("E1").Value = "到職日"
    wsA.Range("F1").Value = "薪資"
    wsA.Range("A2").Value = "（標準格式範例資料列）"

    wsB.Range("A1").Value = "員工編號"
    wsB.Range("B1").Value = "姓名"
    wsB.Range("C1").Value = "職稱"
    wsB.Range("D1").Value = "部門"
    wsB.Range("E1").Value = "薪資"
    wsB.Range("F1").Value = "備註"
    wsB.Range("A2").Value = "（上傳格式範例資料列）"
End Sub

' 比對兩張工作表的欄位標題
Public Sub CompareColumnHeaders(ByVal standardSheet As String, ByVal uploadSheet As String, _
                                 ByVal reportSheet As String)
    Dim wsStd        As Worksheet
    Dim wsUpd        As Worksheet
    Dim wsR          As Worksheet
    Dim stdLastCol   As Long
    Dim updLastCol   As Long
    Dim i            As Long
    Dim j            As Long
    Dim rptRow       As Long
    Dim stdHeader    As String
    Dim updHeader    As String
    Dim foundInUpload As Boolean
    Dim foundInStd   As Boolean
    Dim missingCount As Long
    Dim extraCount   As Long
    Dim orderMismatch As Long

    On Error GoTo ErrHandler

    Set wsStd = ThisWorkbook.Worksheets(standardSheet)
    Set wsUpd = ThisWorkbook.Worksheets(uploadSheet)
    Set wsR = GetOrCreateSheetCCH(reportSheet)

    stdLastCol = wsStd.Cells(1, wsStd.Columns.Count).End(xlToLeft).Column
    updLastCol = wsUpd.Cells(1, wsUpd.Columns.Count).End(xlToLeft).Column

    wsR.Range("A1").Value = "標準格式欄位"
    wsR.Range("B1").Value = "上傳格式欄位"
    wsR.Range("C1").Value = "比對結果"
    With wsR.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(84, 130, 53)
        .Font.Color = RGB(255, 255, 255)
    End With

    rptRow = 2
    missingCount = 0
    extraCount = 0
    orderMismatch = 0

    ' 以標準格式為基準，逐欄比對
    For i = 1 To stdLastCol
        stdHeader = CStr(wsStd.Cells(1, i).Value)
        foundInUpload = False
        For j = 1 To updLastCol
            updHeader = CStr(wsUpd.Cells(1, j).Value)
            If stdHeader = updHeader Then
                foundInUpload = True
                wsR.Cells(rptRow, 1).Value = stdHeader & " (第" & i & "欄)"
                wsR.Cells(rptRow, 2).Value = updHeader & " (第" & j & "欄)"
                If i = j Then
                    wsR.Cells(rptRow, 3).Value = "位置相符"
                Else
                    wsR.Cells(rptRow, 3).Value = "位置不同"
                    wsR.Cells(rptRow, 1).Resize(1, 3).Interior.Color = RGB(255, 235, 156)
                    orderMismatch = orderMismatch + 1
                End If
                rptRow = rptRow + 1
                Exit For
            End If
        Next j
        If Not foundInUpload Then
            wsR.Cells(rptRow, 1).Value = stdHeader & " (第" & i & "欄)"
            wsR.Cells(rptRow, 2).Value = "(缺少)"
            wsR.Cells(rptRow, 3).Value = "上傳格式缺少此欄"
            wsR.Cells(rptRow, 1).Resize(1, 3).Interior.Color = RGB(255, 199, 206)
            rptRow = rptRow + 1
            missingCount = missingCount + 1
        End If
    Next i

    ' 找出上傳格式多出的欄位
    For j = 1 To updLastCol
        updHeader = CStr(wsUpd.Cells(1, j).Value)
        foundInStd = False
        For i = 1 To stdLastCol
            If updHeader = CStr(wsStd.Cells(1, i).Value) Then
                foundInStd = True
                Exit For
            End If
        Next i
        If Not foundInStd Then
            wsR.Cells(rptRow, 1).Value = "(不在標準格式)"
            wsR.Cells(rptRow, 2).Value = updHeader & " (第" & j & "欄)"
            wsR.Cells(rptRow, 3).Value = "上傳格式多出此欄"
            wsR.Cells(rptRow, 1).Resize(1, 3).Interior.Color = RGB(198, 239, 206)
            rptRow = rptRow + 1
            extraCount = extraCount + 1
        End If
    Next j

    wsR.Columns("A:C").AutoFit
    wsR.Activate
    MsgBox "欄位結構比對完成！" & vbCrLf & _
           "缺少欄位: " & missingCount & " 個" & vbCrLf & _
           "多餘欄位: " & extraCount & " 個" & vbCrLf & _
           "位置不符: " & orderMismatch & " 個", vbInformation, "欄位比對結果"
    Exit Sub

ErrHandler:
    MsgBox "發生錯誤: " & Err.Description, vbCritical, "錯誤"
End Sub

Private Function GetOrCreateSheetCCH(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    ws.Cells.Clear
    Set GetOrCreateSheetCCH = ws
End Function
