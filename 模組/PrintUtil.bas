Attribute VB_Name = "PrintUtil"
Option Explicit
'*************************************************************************************
'專案名稱: 全委帳務處理
'功能描述: 列印工具底層元件
'
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2017/7/20
'
'改版日期:
'改版備註: 2017/9/18 增加列印各種檔案
'                 2017/9/20 增加列印PDF
'*************************************************************************************
#If VBA7 Then
    Declare PtrSafe  Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)    'For 64 Bit Systems
    Declare PtrSafe  Function ShellExecute Lib "shell32.dll" _
            Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
                                   ByVal lpFile As String, ByVal lpParameters As String, _
                                   ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#Else
    Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)    'For 32 Bit Systems
    Declare Function ShellExecute Lib "shell32.dll" _
                                  Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
                                                         ByVal lpFile As String, ByVal lpParameters As String, _
                                                         ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

'列印檔案第一個頁籤(根據密碼 or 不使用密碼)
Function PrintFile(ByVal fileName As String, ByVal pwd As String, ByVal usePwd As Boolean)
    Dim windowName As String

    windowName = FileIOUtility.GetFileNameWithoutFolder(fileName)

    If usePwd = True Then
        Workbooks.Open fileName, Password:=pwd
    Else
        Workbooks.Open fileName
    End If

    Windows(windowName).Activate
    Sheets(1).PrintOut
    Windows(windowName).Close savechanges:=False
    Windows(ThisWorkbook.Name).Activate

End Function

'列印excel檔案
Function PrintExcel(ByVal fileName As String)
    Workbooks.Open fileName:=fileName
    '    If InStr(1, fileName, "代操收入費用報告書") > 0 Or InStr(1, fileName, "代操定存應計息表") > 0 Then
    '        AutoPrintIS
    '    End If
    ExecuteExcel4Macro "PRINT(1,,,1,,,,,,,,2,,,TRUE,,FALSE)"
    ActiveWorkbook.Close savechanges:=False
End Function

'自動調整版面為A4橫印
Function AutoPrintIS()
    AutoFitAllColumnsA
    Dim ws As Worksheet
    Set ws = ActiveSheet
    With ws.PageSetup
        .Zoom = False
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .Orientation = xlLandscape
    End With
    'ws.PrintOut
    Set ws = Nothing
End Function

'調整欄寬
Function AutoFitAllColumnsA()
    ActiveSheet.UsedRange.Columns.AutoFit
End Function

' 取得印表機名稱
Function GetDefaultPrinterName()
    Dim Printer As String
    Printer = Application.ActivePrinter
    Printer = Left$(Printer, InStr(Printer, " on") - 1)

    GetDefaultPrinterName = Printer
End Function

'列印檔案
Function PrintAllKindFile(ByVal fileName As String)
    Dim ReturnVal As Variant
    ReturnVal = ShellExecute(0&, "print", fileName, 0&, 0&, 0&)
End Function
'雙面pdf檔案
Function PrintPdfFile(ByVal adobeExePath As String, ByVal fileName As String)
    Dim cmd As String
    cmd = """" & adobeExePath & """  /n /s /h /t """ & fileName & """"
    Shell (cmd)
End Function

'雙面列印檔案
Function PrintAllKindFileDuplex(ByVal fileName As String, Optional ByVal second As Long)
    Dim waitSecond, myduplex, ret As Long
    Dim printerName As String

    printerName = GetDefaultPrinterName

    If second = 0 Then
        waitSecond = 3000
    Else
        waitSecond = second * 1000
    End If

    myduplex = SetPrinterDuplex.Setduplex(printerName, 0)
    ret = SetPrinterDuplex.Setduplex(printerName, 3)  'set duplex on
    ' print your stuff here, including changing to the correct printer
    PrintAllKindFile (fileName)
    Sleep (waitSecond)
    ret = SetPrinterDuplex.Setduplex(printerName, myduplex)   ' return to original duplex setting

End Function

'雙面列印pdf檔案
Function PrintPdfFileDuplex(ByVal adobeExePath As String, ByVal fileName As String, _
                            Optional ByVal second As Long)
    Dim waitSecond, myduplex, ret As Long
    Dim printerName As String

    printerName = GetDefaultPrinterName

    If second = 0 Then
        waitSecond = 3000
    Else
        waitSecond = second * 1000
    End If

    myduplex = SetPrinterDuplex.Setduplex(printerName, 0)
    ret = SetPrinterDuplex.Setduplex(printerName, 3)  'set duplex on
    ' print your stuff here, including changing to the correct printer
    PrintPdfFile adobeExePath:=adobeExePath, fileName:=fileName
    Sleep (waitSecond)
    ret = SetPrinterDuplex.Setduplex(printerName, myduplex)   ' return to original duplex setting

End Function

'列印html檔案
'記得引用Microsoft Internet Controls
Function PrintHtml(ByVal fileNameOrUrl As String)
    Dim ie, TimeOutWebQuery, TimeOutTime As Variant
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Navigate fileNameOrUrl
    ie.Visible = False
    
    TimeOutWebQuery = 5
    TimeOutTime = DateAdd("s", TimeOutWebQuery, Now)
    Do Until ie.ReadyState = 4
        DoEvents
        If Now > TimeOutTime Then
            ie.Stop
            GoTo ErrorTimeOut
        End If
    Loop

    ie.ExecWB 6, 2
    Application.Wait (Now + TimeValue("0:00:03"))

ErrorTimeOut:

    Set ie = Nothing
End Function
