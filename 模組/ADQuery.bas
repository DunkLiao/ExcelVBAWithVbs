Attribute VB_Name = "ADQuery"
Option Explicit
'*************************************************************************************
'專案名稱: AD帳號查詢
'功能描述: 查詢海外分行LOCALHIRE及台幹的帳號及email
'
'版權所有: 合作金庫商業銀行
'程式撰寫: Dunk
'撰寫日期：2016/11/11
'
'改版日期:
'改版備註:
'*************************************************************************************

Public Function StartQueryUsers()
    '選取的分行代號
    Dim branch As Variant
    Dim departId As String
    Dim departName As String
    branch = VBA.Split(ThisWorkbook.Sheets("查詢功能頁面").Range("B2"), "_")
    If UBound(branch) < 1 Then
        MsgBox "請重新選擇分行"
    End If
    
    departId = branch(1)
    departName = branch(0)
        
    Application.ScreenUpdating = False
    ThisWorkbook.Sheets("查詢功能頁面").Range("E:K").ClearContents
    
    '設定表頭
    Dim currCell As Range
    Set currCell = ThisWorkbook.Sheets("查詢功能頁面").Range("E1")
    currCell.Value = "分行名稱"
    currCell.Next.Value = "分行代號"
    currCell.Next.Next.Value = "TCB帳號"
    currCell.Next.Next.Next.Value = "description"
    currCell.Next.Next.Next.Next.Value = "displayName"
    currCell.Next.Next.Next.Next.Next.Value = "mail"
    currCell.Next.Next.Next.Next.Next.Next.Value = "是否為LOCALHIRE"
    
    GetUserInfoLocal departId:=departId, departName:=departName
    ThisWorkbook.Sheets("查詢功能頁面").Columns("E:K").AutoFit
    
    Application.ScreenUpdating = True
    MsgBox "查詢完畢"
End Function

'查詢localhire
Public Function GetUserInfoLocal(ByVal departId As String, ByVal departName As String)
       On Error Resume Next
       

       Const ADS_SCOPE_SUBTREE = 2
       Dim objConnection As Variant
       Dim objCommand As Variant
       Dim objRecordSet As Variant
       Dim ldapString As String
       
       ldapString = "LDAP://OU=" & departId & ",OU=LOCALHIRE,DC=tcb,DC=com"
       'ldapString = "LDAP://OU=A01447,OU=LOCALHIRE,DC=tcb,DC=com"
       Set objConnection = CreateObject("ADODB.Connection")
       Set objCommand = CreateObject("ADODB.Command")
       objConnection.Provider = "ADsDSOObject"
       objConnection.Open "Active Directory Provider"
       Set objCommand.ActiveConnection = objConnection

       objCommand.Properties("Page Size") = 1000
       objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
       
       'search all users from the domn XXX
       objCommand.CommandText = _
               "SELECT ou,description,department,name,displayName,mail FROM '" & ldapString & "' WHERE objectCategory='user'"
       Set objRecordSet = objCommand.Execute
       
       objRecordSet.MoveFirst

       Dim currCell As Range
       Set currCell = ThisWorkbook.Sheets("查詢功能頁面").Range("E2")

       Do Until objRecordSet.EOF
               currCell.Value = departId
               
               currCell.Next.Value = departName
               currCell.Next.Next.NumberFormatLocal = "@"
               currCell.Next.Next.Value = objRecordSet.Fields("name").Value
               currCell.Next.Next.Next.Value = objRecordSet.Fields("description").Value
               currCell.Next.Next.Next.Next.Value = objRecordSet.Fields("displayName").Value
               currCell.Next.Next.Next.Next.Next.Value = objRecordSet.Fields("mail").Value
               currCell.Next.Next.Next.Next.Next.Next.Value = "Y"
               
               Set currCell = currCell.Offset(1, 0)
               objRecordSet.MoveNext
       Loop
End Function

