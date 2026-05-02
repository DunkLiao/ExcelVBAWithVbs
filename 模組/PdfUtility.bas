Attribute VB_Name = "PdfUtility"
Option Explicit
'*************************************************************************************
'專案名稱: 底層元件
'功能描述: PDF工具
'Reference required: "VBE - Tools - References - Acrobat
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2018/3/14
'
'改版日期:
'改版備註: 2018/4/19 增加將pdf存成其他格式
'
'*************************************************************************************
'Sub test()
''    MergePDFs "D:\基金報表\20180312", _
 '     '                   "Crystal Reports_ G202.pdf", _
 '     '                   "D:\基金報表\20180312\黃金報表.pdf"
'    SplitPdf "D:\基金報表\20180312\MergedFile.pdf"
'End Sub

'合併pdf檔案，檔案清單用逗號分隔
Function MergePDFs(ByVal sourceFolder As String, ByVal fileList As String, ByVal destFile As String)
' ZVI:2013-08-27 http://www.vbaexpress.com/forum/showthread.php?47310-Need-code-to-merge-PDF-files-in-a-folder-using-adobe-acrobat-X
' Reference required: "VBE - Tools - References - Acrobat"

    Dim a As Variant, i As Long, n As Long, ni As Long, p As String
    Dim AcroApp As New Acrobat.AcroApp, PartDocs() As Acrobat.CAcroPDDoc

    If Right(sourceFolder, 1) = "\" Then p = sourceFolder Else p = sourceFolder & "\"
    a = Split(fileList, ",")
    ReDim PartDocs(0 To UBound(a))

    On Error GoTo exit_
    If Len(Dir(destFile)) Then Kill destFile
    For i = 0 To UBound(a)
        ' Check PDF file presence
        If Dir(p & Trim(a(i))) = "" Then
            MsgBox "File not found" & vbLf & p & a(i), vbExclamation, "Canceled"
            Exit For
        End If
        ' Open PDF document
        Set PartDocs(i) = CreateObject("AcroExch.PDDoc")
        PartDocs(i).Open p & Trim(a(i))
        If i Then
            ' Merge PDF to PartDocs(0) document
            ni = PartDocs(i).GetNumPages()
            If Not PartDocs(0).InsertPages(n - 1, PartDocs(i), 0, ni, True) Then
                MsgBox "Cannot insert pages of" & vbLf & p & a(i), vbExclamation, "Canceled"
            End If
            ' Calc the number of pages in the merged document
            n = n + ni
            ' Release the memory
            PartDocs(i).Close
            Set PartDocs(i) = Nothing
        Else
            ' Calc the number of pages in PartDocs(0) document
            n = PartDocs(0).GetNumPages()
        End If
    Next

    If i > UBound(a) Then
        ' Save the merged document to DestFile
        If Not PartDocs(0).Save(PDSaveFull, destFile) Then
            MsgBox "Cannot save the resulting document" & vbLf & destFile, vbExclamation, "Canceled"
        End If
    End If

exit_:

    ' Inform about error/success
    If Err Then
        MsgBox Err.Description, vbCritical, "Error #" & Err.Number
    ElseIf i > UBound(a) Then
        'MsgBox "The resulting file is created:" & vbLf & DestFile, vbInformation, "Done"
    End If

    ' Release the memory
    If Not PartDocs(0) Is Nothing Then PartDocs(0).Close
    Set PartDocs(0) = Nothing

    ' Quit Acrobat application
    AcroApp.Exit
    Set AcroApp = Nothing

End Function

'合併pdf檔案，檔案清單用逗號分隔
Function MergePDFsWithList(ByRef fileList() As Variant, ByVal destFile As String)
' ZVI:2013-08-27 http://www.vbaexpress.com/forum/showthread.php?47310-Need-code-to-merge-PDF-files-in-a-folder-using-adobe-acrobat-X
' Reference required: "VBE - Tools - References - Acrobat"

    Dim i As Long, n As Long, ni As Long
    Dim AcroApp As New Acrobat.AcroApp, PartDocs() As Acrobat.CAcroPDDoc

    ReDim PartDocs(0 To UBound(fileList))

    On Error GoTo exit_
    If Len(Dir(destFile)) Then Kill destFile
    For i = 0 To UBound(fileList)
        ' Check PDF file presence
        If Dir(fileList(i)) = "" Then
            MsgBox "File not found" & vbLf & fileList(i), vbExclamation, "Canceled"
            Exit For
        End If
        ' Open PDF document
        Set PartDocs(i) = CreateObject("AcroExch.PDDoc")
        PartDocs(i).Open fileList(i)
        If i Then
            ' Merge PDF to PartDocs(0) document
            ni = PartDocs(i).GetNumPages()
            If Not PartDocs(0).InsertPages(n - 1, PartDocs(i), 0, ni, True) Then
                MsgBox "Cannot insert pages of" & vbLf & fileList(i), vbExclamation, "Canceled"
            End If
            ' Calc the number of pages in the merged document
            n = n + ni
            ' Release the memory
            PartDocs(i).Close
            Set PartDocs(i) = Nothing
        Else
            ' Calc the number of pages in PartDocs(0) document
            n = PartDocs(0).GetNumPages()
        End If
    Next

    If i > UBound(fileList) Then
        ' Save the merged document to DestFile
        If Not PartDocs(0).Save(PDSaveFull, destFile) Then
            MsgBox "Cannot save the resulting document" & vbLf & destFile, vbExclamation, "Canceled"
        End If
    End If

exit_:

    ' Inform about error/success
    If Err Then
        MsgBox Err.Description, vbCritical, "Error #" & Err.Number
    ElseIf i > UBound(fileList) Then
        'MsgBox "The resulting file is created:" & vbLf & DestFile, vbInformation, "Done"
    End If

    ' Release the memory
    If Not PartDocs(0) Is Nothing Then PartDocs(0).Close
    Set PartDocs(0) = Nothing

    ' Quit Acrobat application
    AcroApp.Exit
    Set AcroApp = Nothing

End Function
'分割pdf檔案，一頁一個檔案
Function SplitPdf(ByVal sourceFileName As String)
    Dim PDDoc As Acrobat.CAcroPDDoc, newPDF As Acrobat.CAcroPDDoc
    Dim PDPage As Acrobat.CAcroPDPage
    Dim thePDF As String, PNum, i As Long
    Dim Result As Variant
    Dim NewName, fixFormat As String
    Dim resider As Integer
    Set PDDoc = CreateObject("AcroExch.pdDoc")
    Result = PDDoc.Open(sourceFileName)
    If Not Result Then
        MsgBox "Can't open file: " & sourceFileName
        Exit Function
    End If

    PNum = PDDoc.GetNumPages

    resider = Len(PNum - 1)
    fixFormat = ""
    For i = 1 To resider
        fixFormat = fixFormat & "0"
    Next

    For i = 0 To PNum - 1
        Set newPDF = CreateObject("AcroExch.pdDoc")
        newPDF.Create
        NewName = Replace(sourceFileName, ".pdf", "_" & Format(CStr(i + 1), fixFormat) & ".pdf")
        newPDF.InsertPages -1, PDDoc, i, 1, 0
        newPDF.Save 1, NewName
        newPDF.Close
        Set newPDF = Nothing
    Next i

End Function

'將pdf存成其他格式
Function SavePDFAsOtherFormat(PDFPath As String, FileExtension As String)
   
    'Saves a PDF file as another format using Adobe Professional.
   
    'By Christos Samaras
    'http://www.myengineeringworld.net
   
    'In order to use the macro you must enable the Acrobat library from VBA editor:
    'Go to Tools -> References -> Adobe Acrobat xx.0 Type Library, where xx depends
    'on your Acrobat Professional version (i.e. 9.0 or 10.0) you have installed to your PC.
   
    'Alternatively you can find it Tools -> References -> Browse and check for the path
    'C:\Program Files\Adobe\Acrobat xx.0\Acrobat\acrobat.tlb
    'where xx is your Acrobat version (i.e. 9.0 or 10.0 etc.).
   
    Dim objAcroApp      As Acrobat.AcroApp
    Dim objAcroAVDoc    As Acrobat.AcroAVDoc
    Dim objAcroPDDoc    As Acrobat.AcroPDDoc
    Dim objJSO          As Object
    Dim boResult        As Boolean
    Dim ExportFormat    As String
    Dim NewFilePath     As String
   
    'Check if the file exists.
    If Dir(PDFPath) = "" Then
        MsgBox "Cannot find the PDF file!" & vbCrLf & "Check the PDF path and retry.", _
                vbCritical, "File Path Error"
        Exit Function
    End If
   
    'Check if the input file is a PDF file.
    If LCase(Right(PDFPath, 3)) <> "pdf" Then
        MsgBox "The input file is not a PDF file!", vbCritical, "File Type Error"
        Exit Function
    End If
   
    'Initialize Acrobat by creating App object.
    Set objAcroApp = CreateObject("AcroExch.App")
   
    'Set AVDoc object.
    Set objAcroAVDoc = CreateObject("AcroExch.AVDoc")
   
    'Open the PDF file.
    boResult = objAcroAVDoc.Open(PDFPath, "")
       
    'Set the PDDoc object.
    Set objAcroPDDoc = objAcroAVDoc.GetPDDoc
   
    'Set the JS Object - Java Script Object.
    Set objJSO = objAcroPDDoc.GetJSObject
   
    'Check the type of conversion.
    Select Case LCase(FileExtension)
        Case "eps": ExportFormat = "com.adobe.acrobat.eps"
        Case "html", "htm": ExportFormat = "com.adobe.acrobat.html"
        Case "jpeg", "jpg", "jpe": ExportFormat = "com.adobe.acrobat.jpeg"
        Case "jpf", "jpx", "jp2", "j2k", "j2c", "jpc": ExportFormat = "com.adobe.acrobat.jp2k"
        Case "docx": ExportFormat = "com.adobe.acrobat.docx"
        Case "doc": ExportFormat = "com.adobe.acrobat.doc"
        Case "png": ExportFormat = "com.adobe.acrobat.png"
        Case "ps": ExportFormat = "com.adobe.acrobat.ps"
        Case "rft": ExportFormat = "com.adobe.acrobat.rft"
        Case "xlsx": ExportFormat = "com.adobe.acrobat.xlsx"
        Case "xls": ExportFormat = "com.adobe.acrobat.spreadsheet"
        Case "txt": ExportFormat = "com.adobe.acrobat.accesstext"
        Case "tiff", "tif": ExportFormat = "com.adobe.acrobat.tiff"
        Case "xml": ExportFormat = "com.adobe.acrobat.xml-1-00"
        Case Else: ExportFormat = "Wrong Input"
    End Select
    
    'Check if the format is correct and there are no errors.
    If ExportFormat <> "Wrong Input" And Err.Number = 0 Then
        
        'Format is correct and no errors.
        
        'Set the path of the new file. Note that Adobe instead of xls uses xml files.
        'That's why here the xls extension changes to xml.
        If LCase(FileExtension) <> "xls" Then
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", "." & LCase(FileExtension))
        Else
            NewFilePath = WorksheetFunction.Substitute(PDFPath, ".pdf", ".xml")
        End If
        
        'Save PDF file to the new format.
        boResult = objJSO.SaveAs(NewFilePath, ExportFormat)
        
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
        
        'Close the Acrobat application.
        boResult = objAcroApp.Exit
        
        'Inform the user that conversion was successfully.
        MsgBox "The PDf file:" & vbNewLine & PDFPath & vbNewLine & vbNewLine & _
        "Was saved as: " & vbNewLine & NewFilePath, vbInformation, "Conversion finished successfully"
         
    Else
       
        'Something went wrong, so close the PDF file and the application.
       
        'Close the PDF file without saving the changes.
        boResult = objAcroAVDoc.Close(True)
       
        'Close the Acrobat application.
        boResult = objAcroApp.Exit
       
        'Inform the user that something went wrong.
        MsgBox "Something went wrong!" & vbNewLine & "The conversion of the following PDF file FAILED:" & _
        vbNewLine & PDFPath, vbInformation, "Conversion failed"

    End If
       
    'Release the objects.
    Set objAcroPDDoc = Nothing
    Set objAcroAVDoc = Nothing
    Set objAcroApp = Nothing
       
End Function
