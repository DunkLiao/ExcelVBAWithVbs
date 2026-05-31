Attribute VB_Name = "SplitSheetByLanguage"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetByLanguage
'功能說明: 依指定欄位的語言代碼（如TW/EN/JP）將資料列切割到不同工作表
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/31
'
'*************************************************************************************

Sub TestSplitByLanguage()
    Dim ws As Worksheet
    Set ws = GetOrCreateLangSheet(ThisWorkbook, "語言切割來源")
    Call FillLanguageSampleData(ws)
    Call SplitSheetByLanguage(ws, 3)
    MsgBox "依語言切割完成！", vbInformation, "完成"
End Sub

Sub SplitSheetByLanguage(ByVal sourceWs As Worksheet, ByVal langCol As Integer)
    Dim lastRow    As Long
    Dim i          As Long
    Dim langCode   As String
    Dim destWs     As Worksheet
    Dim destRow    As Long
    Dim langDict   As Object

    Set langDict = CreateObject("Scripting.Dictionary")

    lastRow = sourceWs.Cells(sourceWs.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "來源資料不足，無法切割。", vbExclamation, "警告"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    For i = 2 To lastRow
        langCode = UCase(Trim(CStr(sourceWs.Cells(i, langCol).Value)))
        If langCode = "" Then langCode = "UNKNOWN"

        If Not langDict.Exists(langCode) Then
            On Error Resume Next
            Set destWs = sourceWs.Parent.Worksheets(langCode)
            On Error GoTo 0
            If destWs Is Nothing Then
                Set destWs = sourceWs.Parent.Worksheets.Add( _
                    After:=sourceWs.Parent.Worksheets(sourceWs.Parent.Worksheets.Count))
                destWs.Name = langCode
            End If
            destWs.Cells.Clear
            sourceWs.Rows(1).Copy Destination:=destWs.Rows(1)
            langDict.Add langCode, 2
        End If

        destRow = langDict(langCode)
        Set destWs = sourceWs.Parent.Worksheets(langCode)
        sourceWs.Rows(i).Copy Destination:=destWs.Rows(destRow)
        langDict(langCode) = destRow + 1
    Next i

    Dim langKey As Variant
    For Each langKey In langDict.Keys
        sourceWs.Parent.Worksheets(langKey).Columns.AutoFit
    Next langKey

    Application.ScreenUpdating = True
End Sub

Private Sub FillLanguageSampleData(ByVal ws As Worksheet)
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("產品名稱", "說明", "語言")
    ws.Range("A2:C2").Value = Array("產品A", "繁體中文說明", "TW")
    ws.Range("A3:C3").Value = Array("Product A", "English description", "EN")
    ws.Range("A4:C4").Value = Array("製品A", "Japanese description", "JP")
    ws.Range("A5:C5").Value = Array("產品B", "繁體中文說明", "TW")
    ws.Range("A6:C6").Value = Array("Product B", "English description", "EN")
    ws.Range("A7:C7").Value = Array("Produit A", "Description francais", "FR")
    ws.Columns("A:C").AutoFit
End Sub

Private Function GetOrCreateLangSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = sheetName
    End If
    Set GetOrCreateLangSheet = ws
End Function
