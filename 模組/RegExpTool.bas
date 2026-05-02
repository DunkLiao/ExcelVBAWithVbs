Attribute VB_Name = "RegExpTool"
Option Explicit
'*************************************************************************************
'專案名稱: 全權委託
'功能描述: 正規表示式處理
' Add VBA reference to "Microsoft VBScript Regular Expressions 5.5"
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2017/12/15
'
'改版日期:
'改版備註:
'
'*************************************************************************************
Sub test()
Dim strTest As String
'strTest = "本期投資損益 -110, 341.26"
strTest = "PAGE:"
Debug.Print "FindPageSep:" & FindPageSep(strTest)
End Sub
'取得符合正規表示式的第一個值
Function GetRegMatchFirstValue(ByVal inputValue As String, ByVal pattern As String)
    Dim regex As Object
    Dim matches, Match, subMatch As Variant
    Set regex = CreateObject("VBScript.RegExp")

    With regex
        .pattern = pattern
        .Global = True
    End With

    If regex.test(inputValue) = True Then
        Set matches = regex.Execute(inputValue)
        GetRegMatchFirstValue = matches(0).Value
    Else
        GetRegMatchFirstValue = ""
    End If
End Function

'金額欄位
Function FindAmt(ByVal inputValue As String)
    Dim pattern As String
    pattern = "-?[0-9]{1,3}(,[0-9]{3})*.[0-9]+"
    FindAmt = GetRegMatchFirstValue(inputValue, pattern)
End Function

'日期欄位
Function FindDate(ByVal inputValue As String)
    Dim pattern As String
    pattern = "[0-9]{2,3}.[0-9]{1,2}.[0-9]{1,2}"
    FindDate = GetRegMatchFirstValue(inputValue, pattern)
End Function

'找尋用印欄位
Function FindMgr(ByVal inputValue As String)
    Dim pattern As String
    pattern = "經辦覆核主管"
    FindMgr = GetRegMatchFirstValue(inputValue, pattern)
End Function

'找尋加
Function FindPlus(ByVal inputValue As String)
    Dim pattern As String
    pattern = "加"
    FindPlus = GetRegMatchFirstValue(inputValue, pattern)
End Function

'找尋基金編號
Function FindFundNo(ByVal inputValue As String)
    Dim pattern As String
    pattern = "[0-9]{4,4}"
    FindFundNo = GetRegMatchFirstValue(inputValue, pattern)
End Function

'找尋分頁
Function FindPageSep(ByVal inputValue As String)
    Dim pattern As String
    pattern = "PAGE:"
    If GetRegMatchFirstValue(inputValue, pattern) <> "" Then
        FindPageSep = True
    Else
        FindPageSep = False
    End If
End Function
