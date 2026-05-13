Attribute VB_Name = "SplitSheetToXML"
Option Explicit
'*************************************************************************************
'模組名稱: SplitSheetToXML
'功能說明: 將工作表依指定欄位分割，每個分組匯出為獨立的 XML 檔案
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/5/13
'
'*************************************************************************************

Public Sub SplitToXMLByCategory()
    On Error GoTo ErrHandler
    Dim ws         As Worksheet
    Dim outputDir  As String
    Dim splitCol   As Long
    Dim lastRow    As Long
    Dim lastCol    As Long
    Dim i          As Long
    Dim c          As Long
    Dim category   As String
    Dim xmlContent As String
    Dim fso        As Object
    Dim ts         As Object
    Dim filePath   As String
    Dim dict       As Object
    Dim hdr        As String
    Dim val        As String
    Dim safeKey    As String
    Dim key        As Variant

    Set ws = ActiveSheet
    splitCol = 1
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "工作表資料不足，請確認至少有標題列與一列資料。", vbExclamation, "提示"
        Exit Sub
    End If
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    outputDir = ThisWorkbook.Path & "\XMLOutput\"
    If Dir(outputDir, vbDirectory) = "" Then MkDir outputDir
    Set fso  = CreateObject("Scripting.FileSystemObject")
    Set dict = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        category = Trim(CStr(ws.Cells(i, splitCol).Value))
        If category <> "" And Not dict.Exists(category) Then
            dict.Add category, category
        End If
    Next i

    For Each key In dict.Keys
        xmlContent = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
        xmlContent = xmlContent & "<Data>" & vbCrLf
        For i = 2 To lastRow
            If Trim(CStr(ws.Cells(i, splitCol).Value)) = CStr(key) Then
                xmlContent = xmlContent & "  <Row>" & vbCrLf
                For c = 1 To lastCol
                    hdr = Trim(ws.Cells(1, c).Value)
                    If hdr = "" Then hdr = "Col" & c
                    val = Trim(CStr(ws.Cells(i, c).Value))
                    xmlContent = xmlContent & "    <" & hdr & ">" & val & "</" & hdr & ">" & vbCrLf
                Next c
                xmlContent = xmlContent & "  </Row>" & vbCrLf
            End If
        Next i
        xmlContent = xmlContent & "</Data>"
        safeKey = CStr(key)
        safeKey = Replace(safeKey, "/", "-")
        safeKey = Replace(safeKey, ":", "-")
        safeKey = Replace(safeKey, "*", "-")
        safeKey = Replace(safeKey, "?", "-")
        filePath = outputDir & safeKey & ".xml"
        Set ts = fso.CreateTextFile(filePath, True, False)
        ts.Write xmlContent
        ts.Close
    Next key
    MsgBox "已將工作表依分類欄位分割為 " & dict.Count & " 個 XML 檔案，儲存於：" & outputDir, vbInformation, "完成"
    Exit Sub
ErrHandler:
    MsgBox "發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub

