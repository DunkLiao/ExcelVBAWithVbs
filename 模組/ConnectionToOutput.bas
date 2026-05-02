Attribute VB_Name = "ConnectionToOutput"
Option Explicit
'*************************************************************************************
'專案名稱: VBA專案
'功能描述: 連線到外部資料源
'
'版權所有: 合作金庫商業銀行
'程式撰寫: Dunk
'撰寫日期：2015/7/28
'
'改版日期: 2015/7/30
'改版備註: 加入SQL帳號驗證
'*************************************************************************************

Sub 開始建立表格資料()
Application.ScreenUpdating = False
    Call 刪除所有Sheet
    Dim resultArray() As String
    Dim i As Integer
    Dim ip As String
    Dim dbname As String
    Dim account As String
    
    ip = ThisWorkbook.Sheets(1).Range("$D$2").Text
    dbname = ThisWorkbook.Sheets(1).Range("$D$3").Text
    account = ThisWorkbook.Sheets(1).Range("$D$4").Text
    resultArray = 讀取第一個頁籤的表格定義到陣列()
        
    For i = LBound(resultArray) To UBound(resultArray) - 1
        Call 建立資料表定義(resultArray(i), ip, dbname, account)
    Next i
    Call 刪除所有連線

Application.ScreenUpdating = True
End Sub

Function 讀取第一個頁籤的表格定義到陣列() As String()

Dim my_name, pre_name As String
Dim i, j As Byte
Dim my_range As Range
Dim resultArray() As String
Dim arrayCount As Integer

arrayCount = 0
ReDim resultArray(65536)

my_name = ""
pre_name = ""

Set my_range = ThisWorkbook.Sheets(1).Range("A2")
my_name = my_range.Text

Do While my_name <> pre_name
    On Error Resume Next
    resultArray(arrayCount) = my_range.Text
    'Debug.Print resultArray(arrayCount)
    Set my_range = my_range.Offset(1, 0)
    arrayCount = arrayCount + 1
    my_name = my_range.Text
Loop

ReDim Preserve resultArray(arrayCount)
讀取第一個頁籤的表格定義到陣列 = resultArray()

End Function

Function 取得連線字串(ByVal ip As String, ByVal dbname As String, ByVal account As String) As Variant
    If account <> "" Then
    取得連線字串 = "OLEDB;Provider=SQLOLEDB.1;Persist Security Info=True;User ID=" & account & ";Data Source=" & ip & ";Use Procedure for Prepare=1;Auto Translate=Tr" _
        & "ue;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=False;Initi" _
        & "al Catalog=" & dbname
        
    Else
    取得連線字串 = "OLEDB;Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Data Source=" & ip & ";Use Procedure for Prepare=1;Auto T" _
        & "ranslate=True;Packet Size=4096;Use Encryption for Data=False;Tag with column collation when possible=" _
        & "False;Initial Catalog=" & dbname
        
    End If
End Function



Function 建立資料表定義(ByVal tableName As String, ByVal ip As String, ByVal dbname As String, ByVal account As String)
    '建立資料sheet
        ThisWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = tableName
        Call 取得表格資料(tableName, 取得連線字串(ip, dbname, account))
End Function


Function 取得表格資料(ByVal tableName As String, ByRef connectString As Variant)
Attribute 取得表格資料.VB_ProcData.VB_Invoke_Func = " \n14"
With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=connectString, Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = SqlStatement.取得資料庫表格比對格式(tableName)
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
    ActiveSheet.ListObjects(ActiveSheet.ListObjects.Count).TableStyle = "TableStyleMedium13"
End Function
Function 顯示連線資訊()
Attribute 顯示連線資訊.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim i As Integer
    For i = 1 To ThisWorkbook.Connections.Count
        MsgBox CStr(ThisWorkbook.Connections.Item(i).Name)
    Next
End Function

Function 刪除所有連線()
Do While ThisWorkbook.Connections.Count > 0
        ThisWorkbook.Connections.Item(ThisWorkbook.Connections.Count).Delete
Loop
End Function

Function 刪除所有Sheet()
Application.DisplayAlerts = False
 Dim i As Integer
 For i = ThisWorkbook.Sheets.Count To 2 Step -1
 ThisWorkbook.Sheets(i).Delete
 Next i

 Application.DisplayAlerts = True
End Function
