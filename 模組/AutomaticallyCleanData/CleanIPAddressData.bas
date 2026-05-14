Attribute VB_Name = "CleanIPAddressData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanIPAddressData
'功能說明: 清理 IP 位址欄位，去除空白、標準化格式並驗證合法性的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/15
'
'*************************************************************************************

' 範例進入點
Sub TestCleanIPAddressData()
    Dim ws As Worksheet
    Set ws = GetOrCreateIPSheet("IP位址清理範例")
    Call FillDirtyIPData(ws)
    Call CleanIPColumn(ws, 2)
    MsgBox "IP 位址清理完成！", vbInformation, "完成"
End Sub

' 清理工作表指定欄的 IP 位址資料
' ws: 目標工作表
' colIndex: IP 位址所在欄號
Sub CleanIPColumn(ByVal ws As Worksheet, ByVal colIndex As Long)
    On Error GoTo ErrorHandler

    Dim lastRow As Long
    Dim r       As Long
    Dim raw     As String
    Dim cleaned As String

    Application.ScreenUpdating = False
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        raw = CStr(ws.Cells(r, colIndex).Value)
        cleaned = NormalizeIPAddress(raw)
        ws.Cells(r, colIndex).Value = cleaned

        If cleaned = "無效IP" Then
            ws.Cells(r, colIndex).Interior.Color = RGB(255, 199, 206)
        Else
            ws.Cells(r, colIndex).Interior.ColorIndex = xlNone
        End If
    Next r

    ws.Columns(colIndex).AutoFit
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "清理 IP 位址時發生錯誤：" & Err.Description, vbExclamation, "錯誤"
End Sub

' 標準化 IP 位址字串，回傳合法 IPv4 或「無效IP」
Private Function NormalizeIPAddress(ByVal raw As String) As String
    Dim cleaned  As String
    Dim parts    As Variant
    Dim i        As Integer
    Dim octet    As Long
    Dim result   As String
    Dim j        As Integer
    Dim ch       As String
    Dim valid    As Boolean
    Dim ipResult As String

    cleaned = Trim(raw)
    cleaned = Replace(cleaned, Chr(12288), "")
    cleaned = Replace(cleaned, " ", "")

    result = ""
    For j = 1 To Len(cleaned)
        ch = Mid(cleaned, j, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            result = result & ch
        End If
    Next j
    cleaned = result

    If cleaned = "" Then
        NormalizeIPAddress = "無效IP"
        Exit Function
    End If

    parts = Split(cleaned, ".")
    If UBound(parts) <> 3 Then
        NormalizeIPAddress = "無效IP"
        Exit Function
    End If

    valid = True
    ipResult = ""
    For i = 0 To 3
        If Not IsNumeric(parts(i)) Then
            valid = False
            Exit For
        End If
        octet = CLng(parts(i))
        If octet < 0 Or octet > 255 Then
            valid = False
            Exit For
        End If
        If i > 0 Then ipResult = ipResult & "."
        ipResult = ipResult & CStr(octet)
    Next i

    If valid Then
        NormalizeIPAddress = ipResult
    Else
        NormalizeIPAddress = "無效IP"
    End If
End Function

' 填入含髒資料的 IP 測試資料
Private Sub FillDirtyIPData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:B1").Value = Array("裝置名稱", "IP位址")
    ws.Range("A2:B2").Value = Array("主機A", "192.168.1.1")
    ws.Range("A3:B3").Value = Array("主機B", " 10. 0.0. 1 ")
    ws.Range("A4:B4").Value = Array("主機C", "256.100.1.1")
    ws.Range("A5:B5").Value = Array("主機D", "172.16.254.1")
    ws.Range("A6:B6").Value = Array("主機E", "abc.def.ghi.jkl")
    ws.Range("A7:B7").Value = Array("主機F", "192.168.100")
    ws.Range("A8:B8").Value = Array("主機G", "10.10.10.10")
    ws.Range("A9:B9").Value = Array("主機H", "203.0.113.5")
    ws.Columns("A:B").AutoFit
End Sub

' 取得或建立工作表
Private Function GetOrCreateIPSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetOrCreateIPSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateIPSheet Is Nothing Then
        Set GetOrCreateIPSheet = ThisWorkbook.Worksheets.Add
        GetOrCreateIPSheet.Name = sheetName
    End If
    GetOrCreateIPSheet.Cells.Clear
End Function
