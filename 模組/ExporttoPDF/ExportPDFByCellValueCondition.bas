Attribute VB_Name = "ExportPDFByCellValueCondition"
Option Explicit
'*************************************************************************************
'模組名稱: ExportPDFByCellValueCondition
'功能說明: 根據儲存格值條件判斷是否匯出PDF，並以儲存格內容命名檔案的示範程式
'
'版權所有: Dunk
'程式設計: Dunk
'撰寫日期: 2026/6/19
'
'*************************************************************************************

' 簡化測試入口
Sub TestExportPDFByCellValueCondition()
    Call ExportPDFByCellValueCondition("D:\Temp\PDFOutput")
End Sub

' 根據儲存格值條件匯出PDF
' outputPath: PDF輸出資料夾路徑
Sub ExportPDFByCellValueCondition(ByVal outputPath As String)
    Dim ws As Worksheet
    Dim sheetName As String
    Dim lastRow As Long
    Dim i As Long
    Dim exportCount As Long
    Dim fileName As String
    Dim status As String
    Dim pdfPath As String
    Dim fso As Object
    Dim invalidChars As Variant
    Dim j As Long
    
    sheetName = "PDF條件匯出"
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = sheetName
    End If
    
    ws.Cells.Clear
    Call FillPDFExportData(ws)
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(outputPath) Then
        fso.CreateFolder outputPath
    End If
    
    exportCount = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' 無效檔案名字元陣列
    invalidChars = Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
    
    For i = 2 To lastRow
        status = CStr(ws.Cells(i, 4).Value)
        
        ' 條件：只有狀態為"已完成"的才匯出
        If status = "已完成" Then
            fileName = CStr(ws.Cells(i, 1).Value) & "_" & _
                       CStr(ws.Cells(i, 2).Value) & ".pdf"
            
            ' 將無效檔名字元替換
            For j = LBound(invalidChars) To UBound(invalidChars)
                fileName = Replace(fileName, CStr(invalidChars(j)), "_")
            Next j
            
            pdfPath = outputPath & "\" & fileName
            
            ' 將該列資料標記到暫時區域並匯出
            ws.Cells(i, 1).Resize(1, 4).Copy
            ws.Range("F1").PasteSpecial xlPasteValues
            
            ' 設定列印範圍並匯出PDF
            ws.PageSetup.PrintArea = "F1:I1"
            ws.ExportAsFixedFormat Type:=xlTypePDF, _
                fileName:=pdfPath, _
                Quality:=xlQualityStandard
            
            exportCount = exportCount + 1
        End If
    Next i
    
    ws.Range("F1:I1").Clear
    
    Set fso = Nothing
    
    MsgBox "條件式PDF匯出完成！" & vbCrLf & _
           "共匯出 " & exportCount & " 個PDF檔案至：" & vbCrLf & _
           outputPath, vbInformation, "完成"
End Sub

' 填入PDF條件匯出示範資料
Private Sub FillPDFExportData(ByVal ws As Worksheet)
    ws.Range("A1").Value = "客戶編號"
    ws.Range("B1").Value = "客戶名稱"
    ws.Range("C1").Value = "金額"
    ws.Range("D1").Value = "狀態"
    
    ws.Range("A2").Value = "C001"
    ws.Range("B2").Value = "大華貿易"
    ws.Range("C2").Value = 15000
    ws.Range("D2").Value = "已完成"
    
    ws.Range("A3").Value = "C002"
    ws.Range("B3").Value = "明遠科技"
    ws.Range("C3").Value = 8500
    ws.Range("D3").Value = "進行中"
    
    ws.Range("A4").Value = "C003"
    ws.Range("B4").Value = "金茂實業"
    ws.Range("C4").Value = 22000
    ws.Range("D4").Value = "已完成"
    
    ws.Range("A5").Value = "C004"
    ws.Range("B5").Value = "永豐物流"
    ws.Range("C5").Value = 12000
    ws.Range("D5").Value = "已完成"
    
    ws.Columns("A:D").AutoFit
End Sub
