Attribute VB_Name = "SaveWorkSheetTool"
Option Explicit
'*************************************************************************************
'專案名稱: 底層元件
'功能描述: 工作頁另存成xls工具
'
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期：2018/3/22
'
'改版日期:
'改版備註: 2024/9/13 增加將檔案另存成xlsx
'
'*************************************************************************************

'將結果另存成workbook
Function SaveSheetToWorkBook(ByVal sourceSheetName As String, ByVal destXlsFileName As String _
                                                              , ByVal valueOnly As Boolean)
    Dim wsDest As Worksheet

    Set wsDest = Sheets(sourceSheetName)

    Dim wb As Workbook
    Set wb = Workbooks.Add

    wsDest.Copy Before:=wb.Sheets(1)

    '是否只複製值
    If valueOnly = True Then
        With wb.Sheets(1).UsedRange
            .Value = .Value
        End With
    End If

    While wb.Sheets.Count <> 1
        wb.Sheets(wb.Sheets.Count).Delete
    Wend


    wb.SaveAs destXlsFileName
    wb.Close

    Set wsDest = Nothing
    Set wb = Nothing
End Function

'將檔案另存成xlsx(只有值)
Function SaveXlSMToNewXlsxWithValueOnly(ByVal SavePath)
    Dim SourceBook As Workbook, DestBook As Workbook, SourceSheet As Worksheet, DestSheet As Worksheet
    Dim i As Integer
    
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    Set SourceBook = ThisWorkbook
    Set DestBook = Workbooks.Add
        
    For i = DestBook.Worksheets.Count To 2 Step -1
        DestBook.Worksheets(i).Delete
    Next i
        
    For i = ThisWorkbook.Sheets.Count To 1 Step -1
       Set SourceSheet = ThisWorkbook.Sheets(i)
        SourceSheet.Cells.Copy
        Set DestSheet = DestBook.Worksheets.Add
        With DestSheet.Range("A1")
            .PasteSpecial xlPasteValues
            .PasteSpecial xlPasteFormats
        End With
        DestSheet.Name = SourceSheet.Name
        
        With DestSheet.Tab
                On Error Resume Next
                .Color = SourceSheet.Tab.ThemeColor
                On Error Resume Next
                .ColorIndex = SourceSheet.Tab.ColorIndex
                On Error Resume Next
                .ThemeColor = SourceSheet.Tab.ThemeColor
                On Error Resume Next
                .TintAndShade = SourceSheet.Tab.TintAndShade
        End With
    Next
    
    DestBook.Worksheets(DestBook.Worksheets.Count).Delete
    On Error Resume Next
    Kill SavePath
    DestBook.SaveAs Filename:=SavePath
    DestBook.Close
    
    Set SourceSheet = Nothing
    Set DestSheet = Nothing
    Set SourceBook = Nothing
    Set DestBook = Nothing
                
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
End Function
