Attribute VB_Name = "ExportPDFAndSendEmail"
Option Explicit

' ============================================================
' 範例：將工作表匯出為 PDF 後，自動以 Outlook 寄送
' 功能：匯出 PDF 至暫存路徑，並建立含附件的 Outlook 郵件草稿
' ============================================================

Sub ExportPDFAndSendEmail()
    Dim ws          As Worksheet
    Dim pdfPath     As String
    Dim recipient   As String
    Dim mailSubject As String
    Dim mailBody    As String
    Dim olApp       As Object
    Dim olMail      As Object

    Set ws = ActiveSheet

    recipient = InputBox("請輸入收件人 Email：", "設定寄送資訊")
    If Trim(recipient) = "" Then
        MsgBox "未輸入收件人，操作取消。", vbInformation, "提示"
        Exit Sub
    End If

    mailSubject = InputBox("請輸入郵件主旨：", "設定寄送資訊", ws.Name & " PDF 報表")
    mailBody = "您好，" & vbCrLf & vbCrLf & _
               "請參閱附件 PDF 報表。" & vbCrLf & vbCrLf & _
               "此郵件由系統自動產生。"

    ' 儲存至系統暫存資料夾
    pdfPath = Environ("TEMP") & "\" & ws.Name & "_" & Format(Now, "yyyymmdd_hhnnss") & ".pdf"

    On Error GoTo ErrHandler

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pdfPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    ' 建立 Outlook 郵件草稿並附加 PDF
    Set olApp  = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)

    With olMail
        .To      = recipient
        .Subject = mailSubject
        .Body    = mailBody
        .Attachments.Add pdfPath
        .Display  ' 顯示草稿視窗；改為 .Send 可直接寄出
    End With

    MsgBox "PDF 已產生並附加至郵件草稿，請確認後手動寄出。", vbInformation, "完成"
    Exit Sub

ErrHandler:
    MsgBox "操作失敗：" & Err.Description, vbCritical, "錯誤"
End Sub
