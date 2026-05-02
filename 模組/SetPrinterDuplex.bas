Attribute VB_Name = "SetPrinterDuplex"
Option Explicit

'myduplex = Setduplex("HP Color LaserJet 2605", 0)
'ret = Setduplex("HP Color LaserJet 2605", 3)  'set duplex on
'' print your stuff here, including changing to the correct printer
'ret = Setduplex("HP Color LaserJet 2605", myduplex)   ' return to original duplex setting

'Dim reg As Variant, oreg As Object, mystr As Variant
' Const HKEY_CURRENT_USER = &H80000001
'
'Set oreg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
'
'oreg.enumvalues HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", mystr, Arr
'For Each reg In mystr
'    oreg.getstringvalue HKEY_CURRENT_USER, "Software\Microsoft\Windows NT\CurrentVersion\Devices", reg, regvalue
'    Debug.Print reg & " on " & Mid(regvalue, InStr(regvalue, ",") + 1)
'Next


   Private Const CCHDEVICENAME = 32
   Private Const CCHFORMNAME = 32
   Private Const PRINTER_ACCESS_ADMINISTER = &H4
   Private Const PRINTER_ACCESS_USE = &H8
   Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
   Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
     PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)

   Private Const DM_MODIFY = 8
   Private Const DM_COPY = 2
   Private Const DM_IN_BUFFER = DM_MODIFY
   Private Const DM_OUT_BUFFER = DM_COPY
   Private Const IDOK = 1
   Private Const GMEM_MOVEABLE = &H2
   Private Const GMEM_ZEROINIT = &H40
   Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
   Private Const vbNullPtr = 0&

   ' Add appripriate Constants for what you want to change
   Private Const DM_DUPLEX = &H1000&
   Private Const DM_ORIENTATION = &H1&
   Private Const DM_COPIES = &H100&
   Private Const DMPAPER_A4 = 9
   Private Const DM_PAPERSIZE = &H2&

   ' Constants for Duplex
   Private Const DMDUP_SIMPLEX = 1
   Private Const DMDUP_VERTICAL = 2
   Private Const DMDUP_HORIZONTAL = 3

   ' Constants for Orientation
   Private Const DMORIENT_PORTRAIT = 1
   Private Const DMORIENT_LANDSCAPE = 2

   Private Type ACL
      AclRevision As Byte
      Sbz1 As Byte
      AclSize As Integer
      AceCount As Integer
      Sbz2 As Integer
   End Type

   Private Type SECURITY_DESCRIPTOR
      Revision As Byte
      Sbz1 As Byte
      Control As Long
      Owner As Long
      Group As Long
      Sacl As Long   ' PACL
      Dacl As Long   ' PACL
   End Type

   Private Type PRINTER_DEFAULTS
      pDatatype As String
      pDevMode As Long
      DesiredAccess As Long
   End Type

   Private Type PRINTER_INFO_2
      pServerName As Long    ' Pointer to a String
      pPrinterName As Long   ' Pointer to a String
      pShareName As Long     ' Pointer to a String
      pPortName As Long      ' Pointer to a String
      pDriverName As Long    ' Pointer to a String
      pComment As Long       ' Pointer to a String
      pLocation As Long      ' Pointer to a String
      pDevMode As Long
      pSepFile As Long       ' Pointer to a String
      pPrintProcessor As Long   ' Pointer to a String
      pDatatype As Long      ' Pointer to a String
      pParameters As Long    ' Pointer to a String
      pSecurityDescriptor As Long
      Attributes As Long
      Priority As Long
      DefaultPriority As Long
      StartTime As Long
      UntilTime As Long
      Status As Long
      cJobs As Long
      AveragePPM As Long
   End Type

   Private Type DEVMODE
      dmDeviceName(1 To CCHDEVICENAME) As Byte ' As String * CCHDEVICENAME
      dmSpecVersion As Integer
      dmDriverVersion As Integer
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
      dmFormName(1 To CCHFORMNAME) As Byte ' As String * CCHFORMNAME
      dmUnusedPadding As Integer
      dmBitsPerPel As Integer
      dmPelsWidth As Long
      dmPelsHeight As Long
      dmDisplayFlags As Long
      dmDisplayFrequency As Long
   End Type

   Private Declare Function OpenPrinter Lib "winspool.drv" Alias _
     "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, _
     pDefault As PRINTER_DEFAULTS) As Long
   Private Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" _
     (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, _
     ByVal cbBuf As Long, pcbNeeded As Long) As Long
   Private Declare Function SetPrinter Lib "winspool.drv" Alias "SetPrinterA" _
     (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, _
     ByVal Command As Long) As Long
   Private Declare Function ClosePrinter Lib "winspool.drv" _
     (ByVal hPrinter As Long) As Long
   Private Declare Function DocumentProperties Lib "winspool.drv" Alias _
     "DocumentPropertiesA" (ByVal hwnd As Long, ByVal hPrinter As Long, _
     ByVal pDeviceName As String, pDevModeOutput As Any, _
     pDevModeInput As Any, ByVal fMode As Long) As Long
   Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
     hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

   Public Function Setduplex(myprinter As String, ByVal dupval As Long) As Long
      Dim hGlobal As Long, hPrinter As Long, dwNeeded As Long
      Dim pi2 As PRINTER_INFO_2, i As Integer, pbuf() As Byte
      Dim dm As DEVMODE
      Dim pd As PRINTER_DEFAULTS
      Dim bFlag As Long, lFlag As Long
      Dim SD As SECURITY_DESCRIPTOR, SDbuff() As Byte
      On Error GoTo ABORT
      If dupval < 0 Or dupval > 3 Then MsgBox "invalid duplex value": Exit Function
      hGlobal = 0: hPrinter = 0: dwNeeded = 0
      pd.pDatatype = "TEXT"
      pd.pDevMode = VarPtr(dm)
      pd.DesiredAccess = PRINTER_ALL_ACCESS
      ' Open printer handle (in Windows NT/2000, you need full-access
      ' because you will eventually use SetPrinter)
      bFlag = OpenPrinter(myprinter, hPrinter, pd)
      If (bFlag = 0) Or (hPrinter = vbNullPtr) Then GoTo ABORT

      '  The first GetPrinter() tells you how big the buffer should be in
      '  order to hold all of PRINTER_INFO_2. Note that this usually returns
      '  as FALSE, which only means that the buffer (the third parameter) was
      '  not filled in. You don't want it filled in here.

      Call GetPrinter(hPrinter, 2, 0, 0, dwNeeded)
      If (dwNeeded = 0) Then GoTo ABORT

      '  Allocate enough space for PRINTER_INFO_2

      ReDim pbuf(dwNeeded)
      For i = 0 To dwNeeded
          pbuf(i) = 0
      Next i

      '  The second GetPrinter() call fills in all the current settings,
      '  so all you need to do is modify what you're interested in.

      bFlag = GetPrinter(hPrinter, 2, pbuf(0), dwNeeded, dwNeeded)
      If bFlag = 0 Then GoTo ABORT
      Call CopyMemory(pi2, pbuf(0), Len(pi2))

      If pi2.pSecurityDescriptor <> vbNullPtr Then
        Call CopyMemory(SD, ByVal pi2.pSecurityDescriptor, Len(SD))
        'ReDim SDbuff(Len(SD))
        'Call CopyMemory(SD, SDbuff(0), Len(SD))
      End If

     '  Set orientation to Landscape mode and Duplex to Horizontal
     '  if the driver supports it.

     If pi2.pDevMode <> vbNullPtr Then
          Call CopyMemory(dm, ByVal pi2.pDevMode, Len(dm))
          If dupval = 0 Then ' return current setting only
                Setduplex = dm.dmDuplex
                ClosePrinter hPrinter
                Exit Function
          End If
          'If dm.dmFields And DM_PAPERSIZE Then
          If dm.dmFields And DM_DUPLEX Then
              ' Change the devmode by first setting dmFields to the
              ' members that will change, using a bitwise Or
              dm.dmFields = DM_DUPLEX 'DM_ORIENTATION Or
              'dm.dmPaperSize = DMPAPER_A4
'              dm.dmOrientation = DMORIENT_LANDSCAPE
              dm.dmDuplex = dupval
              Call CopyMemory(ByVal pi2.pDevMode, dm, Len(dm))

              '  Make sure the driver-dependent part of devmode is updated as
              '  necessary.
              lFlag = DocumentProperties(vbNullPtr, hPrinter, _
                       myprinter, _
                       ByVal pi2.pDevMode, ByVal pi2.pDevMode, _
                       DM_IN_BUFFER Or DM_OUT_BUFFER)
              If lFlag <> IDOK Then GoTo ABORT

              '  Update printer information.
   '            Call CopyMemory(ByVal pi2.pDevMode, dm, Len(dm))
              pi2.pSecurityDescriptor = 0
              Call CopyMemory(pbuf(0), pi2, Len(pi2))

              bFlag = SetPrinter(hPrinter, 2, pbuf(0), 0)
              If bFlag = 0 Then GoTo ABORT
              Setduplex = True
              '  The driver supported the change, but it wasn't allowed due to
              '  some other reason (probably lack of permission).
          Else
              MsgBox "This printer does not support Duplexing"
          End If
      Else
          '  The driver doesn't support changing this.
          GoTo ABORT
      End If

ABORT:
      If (hPrinter <> 0) Then Call ClosePrinter(hPrinter)

     '  Clean up.
   End Function

