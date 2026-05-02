Attribute VB_Name = "ConnectDBF"
Option Explicit
Sub F_Sample027()
   'Microsoft ActiveX Data Objects 2.X Library 設定引用項目
   '測試資料 F_Data.mdb F_Data2003.mdb F_Data.accdb
    Dim myCon      As New ADODB.Connection
    Dim myRst      As ADODB.Recordset
    Dim mySqlStr   As String
    Dim myFileName As String
    Dim i          As Long
    Worksheets.Add
    myFileName = "F_Data.mdb"           '讀入檔案
    'myFileName = "F_Data2003.mdb"      '讀入檔案
    'myFileName = "F_Data.accdb"        '讀入檔案
    '以SQL來指定讀入資料
    mySqlStr = "SELECT * FROM F_Tbl01"
    myCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & ThisWorkbook.Path & "\" & myFileName & ";"
    'myCon.Open "Provider=Microsoft.Ace.OLEDB.12.0;" & _   'accdb檔案用
               '"Data Source=" & ThisWorkbook.Path & "\" & myFileName & ";"
    Set myRst = myCon.Execute(mySqlStr)
    If myRst.EOF Then
        MsgBox "沒有符合條件的資料"
    Else
        With myRst
           '欄名
            For i = 1 To .Fields.Count
                Cells(1, i).Value = .Fields(i - 1).Name
            Next
           '記錄
            Range("A2").CopyFromRecordset myRst
            .Close
        End With
    End If
    myCon.Close
    Set myRst = Nothing             '釋放物件
    Set myCon = Nothing
End Sub

