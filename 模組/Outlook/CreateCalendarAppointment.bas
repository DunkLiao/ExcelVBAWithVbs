Attribute VB_Name = "CreateCalendarAppointment"
Option Explicit
'*************************************************************************************
' 專案名稱: 建立行事曆約會範例
' 功能說明: 透過 Outlook 建立一筆約會並顯示給使用者確認
'*************************************************************************************

Private Const OL_APPOINTMENT_ITEM As Long = 1
Private Const OL_BUSY As Long = 2

Public Sub CreateCalendarAppointmentExample()
    On Error GoTo ErrorHandler

    Dim outlookApp As Object
    Dim appointmentItem As Object
    Dim subjectText As String
    Dim startText As String
    Dim minutesText As String
    Dim startTime As Date
    Dim durationMinutes As Long

    subjectText = InputBox("請輸入約會主旨", "建立約會", "例行會議")
    If Len(Trim$(subjectText)) = 0 Then
        Exit Sub
    End If

    startText = InputBox("請輸入開始時間，例如 2026/05/09 14:00", "建立約會", Format$(Now + 1, "yyyy/mm/dd hh:nn"))
    If Not IsDate(startText) Then
        MsgBox "開始時間格式不正確。", vbExclamation, "建立約會"
        Exit Sub
    End If

    minutesText = InputBox("請輸入會議分鐘數", "建立約會", "60")
    If Not IsNumeric(minutesText) Then
        MsgBox "分鐘數必須是數字。", vbExclamation, "建立約會"
        Exit Sub
    End If

    startTime = CDate(startText)
    durationMinutes = CLng(minutesText)

    Set outlookApp = GetCreateCalendarAppointmentOutlookApp()
    Set appointmentItem = outlookApp.CreateItem(OL_APPOINTMENT_ITEM)

    With appointmentItem
        .Subject = subjectText
        .Start = startTime
        .Duration = durationMinutes
        .BusyStatus = OL_BUSY
        .ReminderSet = True
        .ReminderMinutesBeforeStart = 15
        .Body = "此約會由 VBA 範例建立，請確認內容後儲存。"
        .Display
    End With

CleanExit:
    Set appointmentItem = Nothing
    Set outlookApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "建立約會時發生錯誤：" & Err.Description, vbExclamation, "建立約會"
    Resume CleanExit
End Sub

Private Function GetCreateCalendarAppointmentOutlookApp() As Object
    On Error Resume Next

    Set GetCreateCalendarAppointmentOutlookApp = GetObject(, "Outlook.Application")
    If GetCreateCalendarAppointmentOutlookApp Is Nothing Then
        Set GetCreateCalendarAppointmentOutlookApp = CreateObject("Outlook.Application")
    End If
End Function