Attribute VB_Name = "DraftMeetingFromSelectedMail"
Option Explicit
'*************************************************************************************
' 專案名稱: 由所選郵件建立會議邀請草稿範例
' 功能說明: 讀取所選郵件的寄件者與主旨，建立一封會議邀請草稿
'*************************************************************************************

Private Const OL_MAIL_ITEM As Long = 43
Private Const OL_APPOINTMENT_ITEM As Long = 1
Private Const OL_MEETING As Long = 1
Private Const OL_REQUIRED As Long = 1

Public Sub DraftMeetingFromSelectedMailExample()
    On Error GoTo ErrorHandler

    Dim outlookApp As Object
    Dim selectionItems As Object
    Dim sourceMail As Object
    Dim meetingItem As Object
    Dim recipientItem As Object
    Dim startText As String
    Dim startTime As Date

    startText = InputBox("請輸入會議開始時間，例如 2026/05/09 15:00", "建立會議草稿", Format$(Now + 1, "yyyy/mm/dd 10:00"))
    If Not IsDate(startText) Then
        MsgBox "會議開始時間格式不正確。", vbExclamation, "建立會議草稿"
        Exit Sub
    End If
    startTime = CDate(startText)

    Set outlookApp = GetDraftMeetingFromMailOutlookApp()
    Set selectionItems = outlookApp.ActiveExplorer.Selection

    If selectionItems.Count = 0 Then
        MsgBox "請先選取一封郵件。", vbInformation, "建立會議草稿"
        GoTo CleanExit
    End If

    Set sourceMail = selectionItems.Item(1)
    If sourceMail.Class <> OL_MAIL_ITEM Then
        MsgBox "選取項目不是郵件。", vbInformation, "建立會議草稿"
        GoTo CleanExit
    End If

    Set meetingItem = outlookApp.CreateItem(OL_APPOINTMENT_ITEM)
    With meetingItem
        .MeetingStatus = OL_MEETING
        .Subject = "討論：" & sourceMail.Subject
        .Start = startTime
        .Duration = 30
        .Location = "請填寫會議地點"
        .Body = "此會議草稿依所選郵件建立。" & vbCrLf & vbCrLf & sourceMail.Subject
        Set recipientItem = .Recipients.Add(GetDraftMeetingFromMailAddress(sourceMail))
        recipientItem.Type = OL_REQUIRED
        .Display
    End With

CleanExit:
    Set recipientItem = Nothing
    Set meetingItem = Nothing
    Set sourceMail = Nothing
    Set selectionItems = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "建立會議草稿時發生錯誤：" & Err.Description, vbExclamation, "建立會議草稿"
    Resume CleanExit
End Sub

Private Function GetDraftMeetingFromMailOutlookApp() As Object
    On Error Resume Next

    Set GetDraftMeetingFromMailOutlookApp = GetObject(, "Outlook.Application")
    If GetDraftMeetingFromMailOutlookApp Is Nothing Then
        Set GetDraftMeetingFromMailOutlookApp = CreateObject("Outlook.Application")
    End If
End Function

Private Function GetDraftMeetingFromMailAddress(ByVal mailItem As Object) As String
    On Error Resume Next

    GetDraftMeetingFromMailAddress = mailItem.SenderEmailAddress
    If Len(GetDraftMeetingFromMailAddress) = 0 Then
        GetDraftMeetingFromMailAddress = mailItem.SenderName
    End If
End Function