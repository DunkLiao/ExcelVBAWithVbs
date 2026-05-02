Attribute VB_Name = "EncodingUtil"
Option Explicit
'*************************************************************************************
'專案名稱: VBA專案
'功能描述:
'取得檔案編碼方式
'版權所有: 合作金庫商業銀行
'程式撰寫: Dunk
'撰寫日期：2015/11/11
'
'改版日期:
'改版備註:
'*************************************************************************************
Public Enum Encoding
ANSI
Unicode
UnicodeBigEndian
UTF8
End Enum
'取得編碼名稱
Public Function GetEncodingName(ByVal enc As Encoding) As String
    Select Case enc
        Case ANSI: GetEncodingName = "ANSI"
        Case Unicode: GetEncodingName = "Unicode"
        Case UnicodeBigEndian: GetEncodingName = "UnicodeBigEndian"
        Case UTF8: GetEncodingName = "UTF8"
    End Select
End Function
'從文字檔取得編碼方式
Public Function GetEncoding(FileName As String) As Encoding
Dim fBytes(1) As Byte, freeNum As Integer

freeNum = FreeFile

Open FileName For Binary Access Read As #freeNum
Get #freeNum, , fBytes(0)
Get #freeNum, , fBytes(1)
Close #freeNum

If fBytes(0) = &HFF And fBytes(1) = &HFE Then
    GetEncoding = Encoding.Unicode
ElseIf fBytes(0) = &HFE And fBytes(1) = &HFF Then
    GetEncoding = Encoding.UnicodeBigEndian
ElseIf fBytes(0) = &HEF And fBytes(1) = &HBB Then
    GetEncoding = Encoding.UTF8
Else
    GetEncoding = Encoding.ANSI
End If

End Function
'把檔案另存成UTF8
Public Sub FileToUTF8(FileName As String)
Dim fBytes() As Byte, uniString As String, freeNum As Integer
Dim ADO_Stream As Object

freeNum = FreeFile

ReDim fBytes(FileLen(FileName))
Open FileName For Binary Access Read As #freeNum
Get #freeNum, , fBytes
Close #freeNum

uniString = StrConv(fBytes, vbUnicode)

Set ADO_Stream = CreateObject("ADODB.Stream")
With ADO_Stream
    .Type = 2
    .Mode = 3
    .Charset = "utf-8"
    .Open
    .WriteText uniString
    .SaveToFileFileName , 2
    .Close
End With
Set ADO_Stream = Nothing
End Sub

'把檔案另存成unicode
Public Sub FileToUnicode(FileName As String)
Dim fBytes() As Byte, uniString As String, freeNum As Integer
Dim ADO_Stream As Object

freeNum = FreeFile

ReDim fBytes(FileLen(FileName))
Open FileName For Binary Access Read As #freeNum
Get #freeNum, , fBytes
Close #freeNum

uniString = StrConv(fBytes, vbUnicode)

Set ADO_Stream = CreateObject("ADODB.Stream")
With ADO_Stream
    .Type = 2
    .Mode = 3
    .Charset = "unicode"
    .Open
    .WriteText uniString
    .SaveToFile FileName, 2
    .Close
End With
Set ADO_Stream = Nothing
End Sub

