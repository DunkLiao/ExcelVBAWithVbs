Attribute VB_Name = "WebHtmlFetch"
Option Explicit
'*************************************************************************************
'專案名稱: 網頁處理底層工具
'功能描述:
'https://stackoverflow.com/questions/4998715/does-vba-have-any-built-in-url-decoding
'版權所有: Dunk
'程式撰寫: Dunk
'撰寫日期：2024/8/5
'
'改版日期:
'改版備註:
'*************************************************************************************
'Microsoft Internet Controls設定引用項目
'Microsoft HTML Object Library設定引用項目
Public Function FetUrlsFromPage(ByVal mySiteUrl As String)
    Dim myIE      As InternetExplorer
    Dim myDoc     As MSHTML.HTMLDocument
    Dim myLnk     As MSHTML.HTMLAnchorElement
    Dim myErr     As Long
    Dim i         As Long
    Dim counter   As Integer
    counter = 0
    Dim myTempList(65536) As String
    Dim myLinkList() As String
    
    Set myIE = New InternetExplorer
    With myIE
        .Navigate mySiteUrl
        .Visible = True
        Do While .Busy
        Loop
        Do Until .ReadyState = READYSTATE_COMPLETE
        Loop
        Set myDoc = .Document
    End With
    With myDoc
        If .frames.Length > 0 Then
            For i = 0 To .frames.Length _
            - 1
                On Error Resume Next
                myErr = Err.Number
                On Error GoTo 0
                Err.Clear
                If myErr = 0 Then
                    For Each myLnk In .frames(i).Document.Links
                        myTempList(counter) = myLnk.href
                        counter = counter + 1
                    Next
                End If
            Next
        Else
            For Each myLnk In .Links
                myTempList(counter) = myLnk.href
                counter = counter + 1
            Next
        End If
    End With
    myIE.Quit
    Set myIE = Nothing                   '釋放物件
    Set myDoc = Nothing
    ReDim myLinkList(counter)
    For i = 0 To counter
        myLinkList(i) = myTempList(i)
    Next
    FetUrlsFromPage = myLinkList
End Function

Public Function FetUrlTitlesFromPage(ByVal mySiteUrl As String)
    Dim myIE      As InternetExplorer
    Dim myDoc     As MSHTML.HTMLDocument
    Dim myLnk     As MSHTML.HTMLAnchorElement
    Dim myErr     As Long
    Dim i         As Long
    Dim counter   As Integer
    counter = 0
    Dim myTempList(65536) As String
    Dim myLinkList() As String
    Set myIE = New InternetExplorer 'CreateObject("InternetExplorer.Application")
    With myIE
        .Navigate mySiteUrl
        .Visible = True
        Do While .Busy
        Loop
        Do Until .ReadyState = READYSTATE_COMPLETE
        Loop
        Set myDoc = .Document
    End With
    With myDoc
        If .frames.Length > 0 Then
            For i = 0 To .frames.Length - 1
                On Error Resume Next
                myErr = Err.Number
                On Error GoTo 0
                Err.Clear
                If myErr = 0 Then
                    For Each myLnk In .frames(i).Document.Links
                        myTempList(counter) = myLnk.innerText
                        counter = counter + 1
                    Next
                End If
            Next
        Else
            For Each myLnk In .Links
                myTempList(counter) = myLnk.innerText
                counter = counter + 1
            Next
        End If
    End With
    myIE.Quit
    Set myIE = Nothing                   '釋放物件
    Set myDoc = Nothing
    ReDim myLinkList(counter)
    For i = 0 To counter
        myLinkList(i) = myTempList(i)
    Next
    FetUrlTitlesFromPage = myLinkList
End Function

'抓取網頁內容
Function FetchHtmlContent(ByVal myURL As String, ByVal Worksheet As Worksheet)
     Dim urlConnection As String
     urlConnection = "URL;" & myURL
     Worksheet.Columns.Delete
    '指定讀入表單
    With Worksheet
        With .QueryTables.Add(Connection:=urlConnection, Destination:=Range("A1"))
                         '讀入目標儲存格
            .WebSelectionType = xlEntirePage
            .WebFormatting = xlWebFormattingAll
            .Refresh BackgroundQuery:=False
        End With
    .Cells.EntireColumn.AutoFit
    End With
End Function

'將網址編碼,記得引用Microsoft Internet Controls
Function ENCODEURL(varText As Variant, Optional blnEncode = True)
    Static objHtmlfile As Object
    If objHtmlfile Is Nothing Then
        Set objHtmlfile = CreateObject("htmlfile")
        With objHtmlfile.parentWindow
            .execScript "function encode(s) {return encodeURIComponent(s)}", "jscript"
        End With
    End If
    If blnEncode Then
        ENCODEURL = objHtmlfile.parentWindow.encode(varText)
    End If
End Function

'將網址解碼
Function DECODEURL(varText As Variant, Optional blnEncode = True)
    Static objHtmlfile As Object
    If objHtmlfile Is Nothing Then
        Set objHtmlfile = CreateObject("htmlfile")
        With objHtmlfile.parentWindow
            .execScript "function decode(s) {return decodeURIComponent(s)}", "jscript"
        End With
    End If
    If blnEncode Then
        DECODEURL = objHtmlfile.parentWindow.decode(varText)
    End If
End Function
