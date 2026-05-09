Attribute VB_Name = "FlagSelectedMailsForFollowUp"
Option Explicit
'*************************************************************************************
' 專案名稱: 標記所選郵件後續追蹤範例
' 功能說明: 對 Outlook 目前選取郵件加上追蹤旗標與提醒時間
'*************************************************************************************

Private Const OL_MAIL_ITEM As Long = 43
Private Const OL_MARK_TODAY As Long = 0

Public Sub FlagSelectedMailsForFollowUpExample()
    On Error GoTo ErrorHandler

    Dim outlookApp As Object
    Dim selectionItems As Object
    Dim mailItem As Object
    Dim reminderText As String
    Dim reminderTime As Date
    Dim flaggedCount As Long

    reminderText = InputBox("請輸入提醒時間，例如 2026/05/09 17:00", "郵件追蹤", Format$(Now + 1, "yyyy/mm/dd 09:00"))
    If Not IsDate(reminderText) Then
        MsgBox "提醒時間格式不正確。", vbExclamation, "郵件追蹤"
        Exit Sub
    End If

    reminderTime = CDate(reminderText)

    Set outlookApp = GetFlagSelectedMailsOutlookApp()
    Set selectionItems = outlookApp.ActiveExplorer.Selection

    For Each mailItem In selectionItems
        If mailItem.Class = OL_MAIL_ITEM Then
            mailItem.MarkAsTask OL_MARK_TODAY
            mailItem.TaskSubject = "追蹤：" & mailItem.Subject
            mailItem.ReminderSet = True
            mailItem.ReminderTime = reminderTime
            mailItem.Save
            flaggedCount = flaggedCount + 1
        End If
    Next mailItem

    MsgBox "已標記 " & CStr(flaggedCount) & " 封郵件。", vbInformation, "郵件追蹤"

CleanExit:
    Set mailItem = Nothing
    Set selectionItems = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "標記郵件時發生錯誤：" & Err.Description, vbExclamation, "郵件追蹤"
    Resume CleanExit
End Sub

Private Function GetFlagSelectedMailsOutlookApp() As Object
    On Error Resume Next

    Set GetFlagSelectedMailsOutlookApp = GetObject(, "Outlook.Application")
    If GetFlagSelectedMailsOutlookApp Is Nothing Then
        Set GetFlagSelectedMailsOutlookApp = CreateObject("Outlook.Application")
    End If
End Function