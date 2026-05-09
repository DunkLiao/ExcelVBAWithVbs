Attribute VB_Name = "ExportSelectedMailsToCsv"
Option Explicit
'*************************************************************************************
' 專案名稱: 匯出所選郵件清單範例
' 功能說明: 將 Outlook 目前選取郵件的寄件者、主旨、時間匯出為 CSV
'*************************************************************************************

Private Const OL_MAIL_ITEM As Long = 43

Public Sub ExportSelectedMailsToCsvExample()
    On Error GoTo ErrorHandler

    Dim outputFile As String
    Dim outlookApp As Object
    Dim selectionItems As Object
    Dim mailItem As Object
    Dim fileNumber As Integer
    Dim exportedCount As Long

    outputFile = InputBox("請輸入 CSV 輸出路徑", "匯出郵件清單", "C:\Temp\SelectedMails.csv")
    If Len(Trim$(outputFile)) = 0 Then
        Exit Sub
    End If

    EnsureExportSelectedMailsFolder GetExportSelectedMailsFolder(outputFile)

    Set outlookApp = GetExportSelectedMailsOutlookApp()
    Set selectionItems = outlookApp.ActiveExplorer.Selection

    fileNumber = FreeFile
    Open outputFile For Output As #fileNumber
    Print #fileNumber, "寄件者,寄件信箱,主旨,收件時間,附件數"

    For Each mailItem In selectionItems
        If mailItem.Class = OL_MAIL_ITEM Then
            Print #fileNumber, CsvExportSelectedMailsText(mailItem.SenderName) & "," & _
                CsvExportSelectedMailsText(GetExportSelectedMailsSenderAddress(mailItem)) & "," & _
                CsvExportSelectedMailsText(mailItem.Subject) & "," & _
                CsvExportSelectedMailsText(Format$(mailItem.ReceivedTime, "yyyy/mm/dd hh:nn:ss")) & "," & _
                CStr(mailItem.Attachments.Count)
            exportedCount = exportedCount + 1
        End If
    Next mailItem

    Close #fileNumber
    MsgBox "匯出完成，共匯出 " & CStr(exportedCount) & " 封郵件。", vbInformation, "匯出郵件清單"

CleanExit:
    On Error Resume Next
    If fileNumber <> 0 Then
        Close #fileNumber
    End If
    Set mailItem = Nothing
    Set selectionItems = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "匯出郵件清單時發生錯誤：" & Err.Description, vbExclamation, "匯出郵件清單"
    Resume CleanExit
End Sub

Private Function GetExportSelectedMailsOutlookApp() As Object
    On Error Resume Next

    Set GetExportSelectedMailsOutlookApp = GetObject(, "Outlook.Application")
    If GetExportSelectedMailsOutlookApp Is Nothing Then
        Set GetExportSelectedMailsOutlookApp = CreateObject("Outlook.Application")
    End If
End Function

Private Function GetExportSelectedMailsSenderAddress(ByVal mailItem As Object) As String
    On Error Resume Next

    GetExportSelectedMailsSenderAddress = mailItem.SenderEmailAddress
End Function

Private Function CsvExportSelectedMailsText(ByVal value As String) As String
    CsvExportSelectedMailsText = """" & Replace$(value, """", """""") & """"
End Function

Private Function GetExportSelectedMailsFolder(ByVal filePath As String) As String
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    GetExportSelectedMailsFolder = fso.GetParentFolderName(filePath)
    Set fso = Nothing
End Function

Private Sub EnsureExportSelectedMailsFolder(ByVal folderPath As String)
    Dim fso As Object

    If Len(folderPath) = 0 Then
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
    Set fso = Nothing
End Sub