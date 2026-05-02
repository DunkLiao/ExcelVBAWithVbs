Attribute VB_Name = "SheetUtil"
'工作頁常用功能
''''''''''''''''''

'調整欄寬
Sub AutoFitAllColumns()
 ActiveSheet.UsedRange.Columns.AutoFit
End Sub
'調整欄高
Sub AutoFitAllRows()
 ActiveSheet.UsedRange.Rows.AutoFit
End Sub
'產生工作表索引超連結
Sub ShtAdd()
    Dim mySht As Worksheet
    Dim ShtName As String
    Dim myRng As Range
    Dim i, j As Integer
    Dim myWb As Workbook
    Set myWb = ActiveWorkbook
    Set mySht = myWb.Worksheets.Add(Count:=1, Before:=Sheets(1))
    
    On Error Resume Next
   
    mySht.Name = "總表索引"
    Columns(2).Clear                            '儲存格的清除
    Set myRng = Range("A1")
    Range("A1").Value = "總表索引"            '標題的設定
    
    For Each mySht In Worksheets
        ShtName = mySht.Name
        If mySht.Name = "總表索引" Then
            Sheets("總表索引").Activate
            Columns(2).Clear                          '儲存格的清除
            Set myRng = Range("A1")
            Range("A1").Value = "總表索引"            '標題的設定

        Else
        
        'Cells(Rows.Count, 2).End(xlUp).Offset(1).Value = mySht.Name     '參考 說明()
    '    MsgBox mySht.Name
        Set myRng = myRng.Offset(1)
        myRng.Activate
        ActiveCell.Hyperlinks.Add Anchor:=Selection, SubAddress:=ShtName & "!A1", Address:="", TextToDisplay:=ShtName
    
        End If
    Next
    'MsgBox Rows.Count
    Set mySht = Nothing

End Sub

'保護工作表
'使用保護修改的地方
Sub 保護工作表()

     ' Loop through all sheets in the workbook.
     For i = 1 To Sheets.Count
        Sheets(i).Protect
     Next i

End Sub
Sub 取消保護工作表()
'
' 取消保護工作表 Macro
' CR3-31 在 2011/6/7 錄製的巨集
'

'
     For i = 1 To Sheets.Count
          Sheets(i).Unprotect
       Next i
End Sub


Sub 工作表大量命名()
Dim my_name, pre_name As String
Dim i, j As Byte
Dim my_range As Range
Sheets(1).Name = "namelist"
Sheets("namelist").Activate
Range("A1").Activate
my_name = ""
pre_name = ""
my_name = ActiveCell.Text
Set my_range = ActiveCell

Do While my_name <> pre_name
    my_name = my_name
    Sheets("namelist").Activate
    my_range.Offset(1, 0).Activate
    Set my_range = ActiveCell
    my_name = ActiveCell.Text
On Error Resume Next
j = i + 2
    If my_name <> "" Then
    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets(j).Name = my_name
    End If
i = i + 1
'MsgBox my_name & my_name
Loop
Sheets(1).Name = "namelist"
Sheets("namelist").Activate
Range("A1").Activate

End Sub


Sub 刪除所有名稱()
        Dim nm As Name
        For Each nm In ThisWorkbook.Names
            nm.Delete
        Next nm
End Sub

'取得最後一列數號
Function GetLastRowNum(ByVal sheetName As String, ByVal columnAlpha As String)
    Dim lastRowNum As Integer
    On Error GoTo xls
    'xlsx
    lastRowNum = ThisWorkbook.Sheets(sheetName).Range(columnAlpha & "1048576").End(xlUp).Row
    GetLastRowNum = lastRowNum
    Exit Function
xls:
    lastRowNum = ThisWorkbook.Sheets(sheetName).Range(columnAlpha & "65536").End(xlUp).Row
    GetLastRowNum = lastRowNum
End Function

'複製指定工作頁的範圍到目的工作頁
Function CopyRange(ByVal sourceSheetName As String, ByVal sourceColumnName As String _
                                                    , ByVal sourceRawStart As Integer, ByVal sourceRawEnd As Integer _
                                                                                       , ByVal destSheetName As String, ByVal destColumnName As String _
                                                                                                                        , ByVal destRawStart As Integer)

    ThisWorkbook.Sheets(sourceSheetName).Select
    ThisWorkbook.Sheets(sourceSheetName).Range(sourceColumnName & CStr(sourceRawStart) _
                                               & ":" & sourceColumnName & CStr(sourceRawEnd)).Select
    Selection.Copy
    ThisWorkbook.Sheets(destSheetName).Select
    ThisWorkbook.Sheets(destSheetName).Range(destColumnName & CStr(destRawStart)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                                                    :=False, Transpose:=False

End Function



