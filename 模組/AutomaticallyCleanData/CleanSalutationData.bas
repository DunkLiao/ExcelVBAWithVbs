Attribute VB_Name = "CleanSalutationData"
Option Explicit
'*************************************************************************************
'模組名稱: CleanSalutationData
'功能說明: 自動清理與標準化稱謂資料（Mr./Ms./Dr.等）的範例程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/27
'
'*************************************************************************************

Sub TestCleanSalutationData()
    Call CleanAndStandardizeSalutation("稱謂標準化範例")
End Sub

Sub CleanAndStandardizeSalutation(ByVal sheetName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim rawVal As String
    Dim cleanedVal As String
    Dim changeCount As Integer

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If

    ws.Cells.Clear

    ' 填入未標準化的稱謂資料
    ws.Range("A1").Value = "原始稱謂"
    ws.Range("B1").Value = "標準化稱謂"
    ws.Range("A1:B1").Font.Bold = True

    ws.Range("A2").Value = "MR. CHEN"
    ws.Range("A3").Value = "ms. wang"
    ws.Range("A4").Value = "dr.  Lin"
    ws.Range("A5").Value = "mister Huang"
    ws.Range("A6").Value = "miss  chang"
    ws.Range("A7").Value = "prof.  Yang"
    ws.Range("A8").Value = "Mr Liu"
    ws.Range("A9").Value = "MS Lee"
    ws.Range("A10").Value = "Dr  Wu"

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    changeCount = 0

    For i = 2 To lastRow
        rawVal = ws.Cells(i, 1).Value

        ' 轉大寫以利比對
        Dim upperVal As String
        upperVal = UCase(Trim(rawVal))

        ' 標準化稱謂
        cleanedVal = ""

        ' 去除多餘空白
        rawVal = Application.WorksheetFunction.Trim(rawVal)

        If InStr(1, upperVal, "MR") = 1 Or InStr(1, upperVal, "MISTER") = 1 Then
            cleanedVal = Replace(upperVal, "MISTER", "MR.")
            If InStr(1, cleanedVal, "MR.") = 0 Then
                cleanedVal = Replace(cleanedVal, "MR", "Mr. ")
            End If

        ElseIf InStr(1, upperVal, "MS") = 1 Or InStr(1, upperVal, "MISS") = 1 Then
            cleanedVal = Replace(upperVal, "MISS", "MS.")
            If InStr(1, cleanedVal, "MS.") = 0 Then
                cleanedVal = Replace(cleanedVal, "MS", "Ms. ")
            End If

        ElseIf InStr(1, upperVal, "DR") = 1 Or InStr(1, upperVal, "DOCTOR") = 1 Then
            cleanedVal = Replace(upperVal, "DOCTOR", "DR.")
            If InStr(1, cleanedVal, "DR.") = 0 Then
                cleanedVal = Replace(cleanedVal, "DR", "Dr. ")
            End If

        ElseIf InStr(1, upperVal, "PROF") = 1 Then
            If InStr(1, cleanedVal, "PROF.") = 0 Then
                cleanedVal = Replace(cleanedVal, "PROF", "Prof. ")
            End If

        Else
            cleanedVal = rawVal
        End If

        ' 修正大小寫：稱謂首字母大寫，姓名首字母大寫
        If cleanedVal <> "" Then
            cleanedVal = StrConv(cleanedVal, vbProperCase)
            ' 確保 Mr./Ms./Dr. 格式正確
            cleanedVal = Replace(cleanedVal, "Mr ", "Mr. ")
            cleanedVal = Replace(cleanedVal, "Ms ", "Ms. ")
            cleanedVal = Replace(cleanedVal, "Dr ", "Dr. ")
            cleanedVal = Replace(cleanedVal, "Prof ", "Prof. ")
        End If

        If cleanedVal <> ws.Cells(i, 1).Value Then
            changeCount = changeCount + 1
        End If

        ws.Cells(i, 2).Value = cleanedVal
    Next i

    ws.Columns("A:B").AutoFit

    MsgBox "稱謂標準化完成！共處理 " & (lastRow - 1) & " 筆資料，" & _
           "修正 " & changeCount & " 筆。", vbInformation, "完成"
End Sub
