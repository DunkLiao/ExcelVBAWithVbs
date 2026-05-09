Attribute VB_Name = "CreateContactsFromExcelRows"
Option Explicit
'*************************************************************************************
' 專案名稱: 由 Excel 建立 Outlook 聯絡人範例
' 功能說明: 讀取目前工作表 A:C 欄，建立 Outlook 聯絡人
'*************************************************************************************

Private Const OL_CONTACT_ITEM As Long = 2

Public Sub CreateContactsFromExcelRowsExample()
    On Error GoTo ErrorHandler

    Dim outlookApp As Object
    Dim contactItem As Object
    Dim worksheetObject As Object
    Dim lastRow As Long
    Dim rowIndex As Long
    Dim createdCount As Long
    Dim fullName As String
    Dim emailAddress As String
    Dim phoneNumber As String

    Set worksheetObject = ActiveSheet
    lastRow = worksheetObject.Cells(worksheetObject.Rows.Count, "A").End(-4162).Row

    If lastRow < 2 Then
        MsgBox "目前工作表沒有可建立的聯絡人資料。", vbInformation, "建立聯絡人"
        Exit Sub
    End If

    Set outlookApp = GetCreateContactsFromRowsOutlookApp()

    For rowIndex = 2 To lastRow
        fullName = Trim$(CStr(worksheetObject.Cells(rowIndex, "A").Value))
        emailAddress = Trim$(CStr(worksheetObject.Cells(rowIndex, "B").Value))
        phoneNumber = Trim$(CStr(worksheetObject.Cells(rowIndex, "C").Value))

        If Len(fullName) > 0 And Len(emailAddress) > 0 Then
            Set contactItem = outlookApp.CreateItem(OL_CONTACT_ITEM)
            With contactItem
                .FullName = fullName
                .Email1Address = emailAddress
                .BusinessTelephoneNumber = phoneNumber
                .Save
            End With
            createdCount = createdCount + 1
        End If
    Next rowIndex

    MsgBox "聯絡人建立完成，共建立 " & CStr(createdCount) & " 筆。", vbInformation, "建立聯絡人"

CleanExit:
    Set contactItem = Nothing
    Set worksheetObject = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "建立聯絡人時發生錯誤：" & Err.Description, vbExclamation, "建立聯絡人"
    Resume CleanExit
End Sub

Private Function GetCreateContactsFromRowsOutlookApp() As Object
    On Error Resume Next

    Set GetCreateContactsFromRowsOutlookApp = GetObject(, "Outlook.Application")
    If GetCreateContactsFromRowsOutlookApp Is Nothing Then
        Set GetCreateContactsFromRowsOutlookApp = CreateObject("Outlook.Application")
    End If
End Function