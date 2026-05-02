Attribute VB_Name = "SettingPrinterDuplexNew"
Option Explicit
'*************************************************************************************
'專案名稱: 底層元件
'功能描述: 設定印表機雙面列印
'
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2017/7/25
'
'改版日期:
'改版備註:
'
'*************************************************************************************

#If VBA7 Then
Public Type PRINTER_INFO_9
    pDevmode As LongPtr    '''' POINTER TO DEVMODE
End Type
#Else
Public Type PRINTER_INFO_9
    pDevmode As Long    '''' POINTER TO DEVMODE
End Type
#End If

Public Type DEVMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer: dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
End Type
#If VBA7 Then
    Public Declare PtrSafe Function GetProfileStringA Lib "kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As LongPtr) As Long

    Public Declare PtrSafe Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As LongPtr, pDefault As Any) As Long
    Public Declare PtrSafe Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As LongPtr, ByVal Level As LongPtr, buffer As LongPtr, ByVal pbSize As LongPtr, pbSizeNeeded As LongPtr) As Long
    Public Declare PtrSafe Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As LongPtr, ByVal Level As LongPtr, pPrinter As Any, ByVal Command As LongPtr) As Long
    Public Declare PtrSafe Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As LongPtr, ByVal hPrinter As LongPtr, ByVal pDeviceName As String, _
                                                                                                       ByVal pDevModeOutput As LongPtr, ByVal pDevModeInput As LongPtr, _
                                                                                                       ByVal fMode As LongPtr) As Long
    Public Declare PtrSafe Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As LongPtr) As Long
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cbLength As LongPtr)

#Else
    Public Declare Function GetProfileStringA Lib "kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

    Public Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Long
    Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, buffer As Long, ByVal pbSize As Long, pbSizeNeeded As Long) As Long
    Public Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal Command As Long) As Long
    Public Declare Function DocumentProperties Lib "winspool.drv" Alias "DocumentPropertiesA" (ByVal hWnd As Long, ByVal hPrinter As Long, ByVal pDeviceName As String, _
                                                                                               ByVal pDevModeOutput As Long, ByVal pDevModeInput As Long, _
                                                                                               ByVal fMode As Long) As Long
    Public Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cbLength As Long)
#End If
Public Const DM_IN_BUFFER = 8
Public Const DM_OUT_BUFFER = 2

Public Sub SetPrinterProperty(ByVal sPrinterName As String, ByVal iPropertyType As Long)
    Dim PrinterName, sPrinter, sDefaultPrinter As String
    Dim Pinfo9 As PRINTER_INFO_9
    Dim hPrinter, nRet As Long
    Dim yDevModeData() As Byte
    Dim dm As DEVMODE

    '''' STROE THE CURRENT DEFAULT PRINTER
    sDefaultPrinter = sPrinterName

    '''' USE THE FULL PRINTER ADDRESS TO GET THE ADDRESS AND NAME MINUS THE PORT NAME
    PrinterName = Left(sDefaultPrinter, InStr(sDefaultPrinter, " on ") - 1)

    '''' OPEN THE PRINTER
    nRet = OpenPrinter(PrinterName, hPrinter, ByVal 0&)

    '''' GET THE SIZE OF THE CURRENT DEVMODE STRUCTURE
    nRet = DocumentProperties(0, hPrinter, PrinterName, 0, 0, 0)
    If (nRet < 0) Then MsgBox "Cannot get the size of the DEVMODE structure.": Exit Sub

    '''' GET THE CURRENT DEVMODE STRUCTURE
    ReDim yDevModeData(nRet + 100) As Byte
    nRet = DocumentProperties(0, hPrinter, PrinterName, VarPtr(yDevModeData(0)), 0, DM_OUT_BUFFER)
    If (nRet < 0) Then MsgBox "Cannot get the DEVMODE structure.": Exit Sub

    '''' COPY THE CURRENT DEVMODE STRUCTURE
    Call CopyMemory(dm, yDevModeData(0), Len(dm))

    '''' CHANGE THE DEVMODE STRUCTURE TO REQUIRED
    dm.dmDuplex = iPropertyType    ' 1 = simplex, 2 = duplex

    '''' REPLACE THE CURRENT DEVMODE STRUCTURE WITH THE NEWLEY EDITED
    Call CopyMemory(yDevModeData(0), dm, Len(dm))

    '''' VERIFY THE NEW DEVMODE STRUCTURE
    nRet = DocumentProperties(0, hPrinter, PrinterName, VarPtr(yDevModeData(0)), VarPtr(yDevModeData(0)), DM_IN_BUFFER Or DM_OUT_BUFFER)

    Pinfo9.pDevmode = VarPtr(yDevModeData(0))

    '''' SET THE DEMODE STRUCTURE WITH ANY CHANGES MADE
    nRet = SetPrinter(hPrinter, 9, Pinfo9, 0)
    If (nRet <= 0) Then MsgBox "Cannot set the DEVMODE structure.": Exit Sub

    '''' CLOSE THE PRINTER
    nRet = ClosePrinter(hPrinter)

End Sub
'設定預設印表機雙面長邊
Public Function SetPrinterDuplexLongSide()
    Dim sPrinterName As String
    sPrinterName = GetDefaultPrinter
    SetPrinterProperty sPrinterName, 2
End Function
'設定預設印表機雙面短邊
Public Function SetPrinterDuplexShortSide()
    Dim sPrinterName As String
    sPrinterName = GetDefaultPrinter
    SetPrinterProperty sPrinterName, 3
End Function
'設定預設印表機單面
Public Function SetPrinterSimplex()
    Dim sPrinterName As String
    sPrinterName = GetDefaultPrinter
    SetPrinterProperty sPrinterName, 1
End Function

'取得預設印表機
Function GetDefaultPrinter()
    Dim Printer As String
    Printer = Application.ActivePrinter
    GetDefaultPrinter = Printer
End Function

'取得印表機資訊
Function DefaultPrinterInfo(ByVal itemName As String)
    Dim strLPT As String * 255
    Dim Result, Printer, Driver, Port As String
    Dim ResultLength As Long
    Dim Comma1, Comma2 As Variant
    Call GetProfileStringA("Windows", "Device", "", strLPT, 254)

    Result = Application.Trim(strLPT)
    ResultLength = Len(Result)

    Comma1 = Application.Find(",", Result, 1)
    Comma2 = Application.Find(",", Result, Comma1 + 1)

    '   Gets printer's name
    Printer = Left(Result, Comma1 - 1)

    '   Gets driver
    Driver = Mid(Result, Comma1 + 1, Comma2 - Comma1 - 1)

    '   Gets last part of device line
    Port = Right(Result, ResultLength - Comma2)

    Select Case itemName
    Case "Printer":
        DefaultPrinterInfo = Printer
        Exit Function
    Case "Driver":
        DefaultPrinterInfo = Driver
        Exit Function
    Case "Port":
        DefaultPrinterInfo = Port
        Exit Function
    End Select
End Function



