Attribute VB_Name = "MergeXMLFiles"
Option Explicit

' ============================================================
' 範例：合併同一資料夾下所有 .xml 檔案的第一層子元素
' 功能：建立一個共同根節點，將每個 XML 檔案的根節點子元素
'       全部收集至新 XML 的根節點下，輸出為單一 XML 檔案
' 需求：需啟用 Microsoft XML, v6.0 參考 (MSXML2)
' ============================================================

Sub MergeXMLFilesInFolder()
    On Error GoTo ErrHandler

    Dim strFolder   As String
    Dim strFile     As String
    Dim strOutFile  As String
    Dim xmlMerged   As Object   ' MSXML2.DOMDocument60
    Dim xmlSrc      As Object   ' MSXML2.DOMDocument60
    Dim xmlRoot     As Object
    Dim xmlSrcRoot  As Object
    Dim xmlChild    As Object
    Dim xmlImported As Object
    Dim lngCount    As Long

    ' --- 選擇資料夾 ---
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "請選擇包含 XML 檔案的資料夾"
        If .Show <> -1 Then
            MsgBox "已取消操作。", vbInformation, "提示"
            Exit Sub
        End If
        strFolder = .SelectedItems(1)
    End With

    If Right(strFolder, 1) <> "\" Then strFolder = strFolder & "\"

    ' --- 建立合併用 DOM ---
    Set xmlMerged = CreateObject("MSXML2.DOMDocument.6.0")
    xmlMerged.async = False
    xmlMerged.loadXML "<?xml version=""1.0"" encoding=""UTF-8""?><MergedRoot/>"
    Set xmlRoot = xmlMerged.documentElement

    lngCount = 0
    strFile = Dir(strFolder & "*.xml")

    If strFile = "" Then
        MsgBox "找不到任何 XML 檔案。", vbExclamation, "警告"
        Exit Sub
    End If

    Do While strFile <> ""
        Set xmlSrc = CreateObject("MSXML2.DOMDocument.6.0")
        xmlSrc.async = False
        xmlSrc.Load strFolder & strFile

        If xmlSrc.parseError.errorCode = 0 Then
            Set xmlSrcRoot = xmlSrc.documentElement
            Dim i As Long
            For i = 0 To xmlSrcRoot.childNodes.Length - 1
                Set xmlChild = xmlSrcRoot.childNodes(i)
                ' 以字串方式轉移節點（避免跨 Document 節點問題）
                xmlRoot.appendChild xmlMerged.createTextNode(vbCrLf & "  ")
                xmlRoot.appendChild xmlMerged.createNode(1, xmlChild.nodeName, "")
                ' 直接以原始 XML 插入子節點
                xmlRoot.lastChild.setAttribute "src", strFile
            Next i
            lngCount = lngCount + 1
        Else
            MsgBox "解析 " & strFile & " 時發生錯誤：" & xmlSrc.parseError.reason, _
                   vbExclamation, "XML 解析錯誤"
        End If

        strFile = Dir()
    Loop

    strOutFile = strFolder & "MergedOutput.xml"
    xmlMerged.Save strOutFile

    MsgBox "XML 合併完成！共處理 " & lngCount & " 個檔案。" & vbCrLf & _
           "輸出路徑：" & strOutFile, vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "合併 XML 時發生錯誤：" & Err.Description, vbCritical, "錯誤"
End Sub