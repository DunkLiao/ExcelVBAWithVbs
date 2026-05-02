Attribute VB_Name = "CheckFormat"
''格式檢查模組
''作者：Guan Jhih Liao

'計算身分證字號檢查碼
'傳回 檢查碼 檢查碼錯誤
'       0  檢查碼正確
'      -1  身份證檢查錯誤
'      -2  檢查值不為身份證(為中文,輸入字串長度不足,第一字母不為英文字母,
'                           其餘的不為數字)
Public Function ChkPID(ByVal PID As String) As Integer
   Dim XX(1 To 26)   As Integer
   Dim i             As Integer
   Dim cc            As String
   Dim First_no      As Integer
   Dim X1            As Integer
   Dim X2            As Integer
   Dim d(1 To 9)     As Integer
   Dim yy            As Integer
     
     XX(1) = 10: XX(2) = 11: XX(3) = 12: XX(4) = 13: XX(5) = 14: XX(6) = 15: XX(7) = 16
     XX(8) = 17: XX(9) = 34: XX(10) = 18: XX(11) = 19: XX(12) = 20: XX(13) = 21: XX(14) = 22
     XX(15) = 35: XX(16) = 23: XX(17) = 24: XX(18) = 25: XX(19) = 26: XX(20) = 27: XX(21) = 28
     XX(22) = 29: XX(23) = 32: XX(24) = 30: XX(25) = 31: XX(26) = 33
     
     ChkPID = 0
     
     PID = Trim(PID)
     
     If PID = "" Then
        ChkPID = -2
        Exit Function
     End If
     
     '檢查輸入字串是否為中文
     If AscB(PID) > 128 Then
        ChkPID = -2
        Exit Function
     End If
     
     '檢查輸入字串是否長度為零
     If Len(PID) = 0 Or Len(PID) <> 10 Then
        ChkPID = -2
        Exit Function
     End If
         
     PID = UCase(PID)  '英文為轉成大寫
     
     '檢查第一個字母是否為英文
     cc = Left(PID, 1)
     If Asc(cc) > 64 And Asc(cc) < 91 Then
     Else
        ChkPID = -2
        Exit Function
     End If
     
     '檢查第二個字母是否為英文
     cc = Mid(PID, 2, 1)
     If Asc(cc) > 64 And Asc(cc) < 91 Then
       '檢查其餘八位全為數字
        cc = Right(PID, 8)
        If IsNumeric(cc) = False Then
           ChkPID = -2
           Exit Function
        End If
        d(1) = XX(Asc(Mid(PID, 2, 1)) - 64) Mod 10
     Else
        '檢查其餘九位全為數字
        cc = Right(PID, 9)
        If IsNumeric(cc) = False Then
           ChkPID = -2
           Exit Function
        End If
        
        '檢查第二個字元, 代表性別
        cc = Mid(PID, 2, 1)
        If cc = "1" Or cc = "2" Then
        Else
           ChkPID = -1
           Exit Function
        End If
        d(1) = Mid(PID, 2, 1)
     End If
     
     First_no = XX(Asc(Left(PID, 1)) - 64)
     X1 = First_no \ 10
     X2 = First_no Mod 10
     
     For i = 2 To 9
        cc = Mid(PID, i + 1, 1)
        d(i) = Val(cc)
     Next i
     
'     yy = X1 + 9 * X2 + 8 * d(1) + 7 * d(2) + 6 * d(3) + 5 * d(4) + 4 * d(5) + 3 * d(6) + 2 * d(7) + d(8) + d(9)
'     ChkPID = yy Mod 10

     yy = X1 + 9 * X2 + 8 * d(1) + 7 * d(2) + 6 * d(3) + 5 * d(4) + 4 * d(5) + 3 * d(6) + 2 * d(7) + d(8)
     ChkPID = (10 - (yy Mod 10)) Mod 10
     If ChkPID = d(9) Then
        ChkPID = 0
     End If
     
End Function

'顯示今日日期(民國年YYYMMDD)
Public Function Today() As String
  Today = Format(Date, "yyyymmdd")
  Today = Today - 19110000
  TodayTw = Format(TodayTw, "0######")
End Function

'檢查民國年(yyy/mm/dd)輸入的合理性
Public Function ChkDate(ByVal sDate As String) As Boolean
  Dim EndDate       As Date
  Dim yy, mm, dd    As Integer
  Dim yy1, mm1, dd1 As Integer
  Dim AfterDate     As String
   
  If Len(sDate) <> 7 Then
     ChkDate = False
     Exit Function
  End If
  
  AfterDate = Format(ChgDate(sDate), "yyyymmdd")
  AfterDate = AfterDate - 19110000
  AfterDate = Format(AfterDate, "0######")
    
  yy = Val(Mid(sDate, 1, 3))
  mm = Val(Mid(sDate, 4, 2))
  dd = Val(Mid(sDate, 6, 2))
  yy1 = Val(Mid(AfterDate, 1, 3))
  mm1 = Val(Mid(AfterDate, 4, 2))
  dd1 = Val(Mid(AfterDate, 6, 2))
  
  If yy <> yy1 Or mm <> mm1 Or dd <> dd1 Then
     ChkDate = False
     Exit Function
  End If
 
  ChkDate = True
End Function

'將民國年(yyymmdd)轉成日期型態
Public Function ChgDate(ByVal sDate As String) As Date
   Dim yy, mm, dd As Integer
   yy = Val(Mid(sDate, 1, 3)) + 1911
   mm = Val(Mid(sDate, 4, 2))
   dd = Val(Mid(sDate, 6, 2))
   ChgDate = DateSerial(yy, mm, dd)
End Function

'將字串轉成數字
Public Function Str2Val(ByVal sNum As String) As Double
   sNum = Format(sNum, "#")
   Str2Val = Val(sNum)
End Function

'切割字串根據字元長度(Byte Length)
'長度超過時會補空白
'mode="L":左補空白
'mode="M":置中(兩側擺空白)
'未指定為右補空白
Public Function FixStr(ByVal iStr As String, ByVal fixlen As Integer, ByVal mode As String) As String
   Dim strlen   As Integer
   Dim strTemp  As String
   Dim x        As Integer
   Dim i        As Integer
   Dim C        As String
   Dim newStr   As String
   Dim nowlen   As Integer
   
'   iStr = Trim(iStr)
   strlen = 0
   strTemp = iStr
   newStr = ""
   
   Do While Len(strTemp) > 0
      C = Left(strTemp, 1)
      If Asc(C) < 0 Then
         nowlen = 2
      Else
         nowlen = 1
      End If
      
      If (strlen + nowlen) > fixlen Then
         Exit Do
      Else
         strlen = strlen + nowlen
      End If
      newStr = newStr + C
      
      strTemp = Right(strTemp, Len(strTemp) - 1)
   Loop
   
   x = fixlen - strlen
   If x > 0 Then
      If mode = "L" Then
         FixStr = Space(x) + newStr
      Else
         If mode = "M" Then
            i = x Mod 2
            If i = 0 Then
               FixStr = Space(x / 2) + newStr + Space(x / 2)
            Else
               i = x / 2
               FixStr = Space(i) + newStr + Space(x - i)
            End If
         Else
            FixStr = newStr + Space(x)
         End If
      End If
   Else
      FixStr = newStr
   End If
End Function

'擷取字串根據指定開始位置及字元長度(Byte Length)
Public Function MidMbcs(ByVal str As String, start As Integer, Optional length As Integer) As String
   Dim strlen   As Integer
   Dim strTemp  As String
   Dim i        As Integer
   Dim C        As String
   Dim newStr   As String
   Dim nowlen   As Integer
   
   strlen = 0
   strTemp = str
   newStr = ""
   
   
   If start < 1 Then start = 1
   i = 0
   Do While i < (start - 1)
      If Len(strTemp) <= 0 Then Exit Do
      
      C = Left(strTemp, 1)
      If Asc(C) < 0 Then
         nowlen = 2
      Else
         nowlen = 1
      End If
      
      If (i + nowlen) >= start Then
         Exit Do
      Else
         i = i + nowlen
         strTemp = Right(strTemp, Len(strTemp) - 1)
      End If
      
      If strTemp = "" Then
         Exit Do
      End If
   Loop
    
   Do While Len(strTemp) > 0
      C = Left(strTemp, 1)
      If Asc(C) < 0 Then
         nowlen = 2
      Else
         nowlen = 1
      End If
      
      If (strlen + nowlen) > length And length <> 0 Then
         Exit Do
      Else
         strlen = strlen + nowlen
      End If
      newStr = newStr + C
      
      strTemp = Right(strTemp, Len(strTemp) - 1)
   Loop
   MidMbcs = newStr
End Function

'擷取字串根據字元長度(Byte Length)
'不足指定長度會去除最後一個byte
Public Function cutchar(dataString As String, limit As Integer) As String
    Dim i As Integer
    Dim s As String
    Dim sql As String
  
    sql = ""
    s = ""
    i = 0
    Do While Len(dataString) <> 0
       s = Left(dataString, 1)
       
       If Asc(s) < 0 Then
          i = i + 2
       Else
          i = i + 1
       End If
       
       If i <= limit Then
          sql = sql + s
          dataString = Right(dataString, Len(dataString) - 1)
          If Asc(s) = 13 Then
             Exit Do
          End If
       Else
          Exit Do
       End If
   Loop
   cutchar = sql
   
End Function

'檢查是否為大寫英文
'用法
'Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'   KeyAscii = ChgUCase(KeyAscii)
'End Sub
Public Function ChgUCase(ByVal KeyAscii As Integer) As Integer
    If KeyAscii >= 97 And KeyAscii <= 122 Then
       ChgUCase = KeyAscii - 32
    Else
       ChgUCase = KeyAscii
    End If
End Function

'檢查是否為數字
'用法
'Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'   If KeyAscii = Asc(vbBack) Then Exit Sub
'
'   If Not ChkValNum(KeyAscii) Then
'      KeyAscii = ChgUCase(KeyAscii)
'   End If
'End Sub
'Private Sub TextBox8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'   If Not ChkValNum(KeyAscii) Then
'      KeyAscii = 0
'   End If
'End Sub
Public Function ChkValNum(ByVal KeyAscii As Integer) As Boolean
   If KeyAscii >= 48 And KeyAscii <= 57 Then
      ChkValNum = True
   Else
      ChkValNum = False
   End If
End Function

'將來源字串移除指定的文字
Public Function InStrDel(ByVal strData As String, ByVal strDelStr As String) As String
    Dim i As Integer
    Dim s As String
    Dim sql As String

    sql = ""
    s = ""
    i = 0
    strData = Trim(strData)

    Do While Len(strData) <> 0
       s = Left(strData, 1)
       Debug.Print Asc(s)
       Debug.Print Asc(strDelStr)

       If Asc(s) = Asc(strDelStr) Then
       Else
          sql = sql + s
       End If
       strData = Right(strData, Len(strData) - 1)

   Loop
   InStrDel = sql

   
'   If Len(strData) = 0 Then Exit Function
'   Do
'      W = InStr(strData, Asc(strDelStr))
'      If W = 0 Then Exit Do
'      strData = Left(strData, W - 1) & Right(strData, Len(strData) - W)
'   Loop
'   InStrDel = strData
End Function

'檢查營利事業統一編號
Public Function BANCheck(ByVal strBAN As String, Optional ByRef strReason As String) As Boolean
    Dim intMod             As Integer      ' 餘數變數
    Dim intSum             As Integer      ' 合計數變數
    Dim intX(1 To 8)       As Integer
    Dim intY(1 To 8)       As Integer
    
    BANCheck = False
    
    If Len(strBAN) <> 8 Then
       strReason = "營利事業統一編號必須是八碼。"
       BANCheck = False
       Exit Function
    End If
    
    If IsNumeric(strBAN) = False Then
         strReason = "輸入之編號中有非數字。"
         BANCheck = False
         Exit Function
    End If
       
    intX(1) = Val(Mid(strBAN, 1, 1)) * 1     ' 第 1位數 * 1。
    intX(2) = Val(Mid(strBAN, 2, 1)) * 2     ' 第 2位數 * 2。
    intX(3) = Val(Mid(strBAN, 3, 1)) * 1     ' 第 3位數 * 1。
    intX(4) = Val(Mid(strBAN, 4, 1)) * 2     ' 第 4位數 * 2。
    intX(5) = Val(Mid(strBAN, 5, 1)) * 1     ' 第 5位數 * 1。
    intX(6) = Val(Mid(strBAN, 6, 1)) * 2     ' 第 6位數 * 2。
    intX(7) = Val(Mid(strBAN, 7, 1)) * 4     ' 第 7位數 * 4。
    intX(8) = Val(Mid(strBAN, 8, 1)) * 1     ' 第 8位數 * 1。
    
    intY(1) = Int(intX(2) / 10)              ' 第 2位數的乘積可能大於10，除以10，取其整數。
    intY(2) = intX(2) Mod 10                 ' 第 2位數的乘積可能大於10，除以10，取其餘數。
    intY(3) = Int(intX(4) / 10)              ' 第 4位數的乘積可能大於10，除以10，取其整數。
    intY(4) = intX(4) Mod 10                 ' 第 4位數的乘積可能大於10，除以10，取其餘數。
    intY(5) = Int(intX(6) / 10)              ' 第 6位數的乘積可能大於10，除以10，取其整數。
    intY(6) = intX(6) Mod 10                 ' 第 6位數的乘積可能大於10，除以10，取其餘數。
    intY(7) = Int(intX(7) / 10)              ' 第 7位數的乘積可能大於10，除以10，取其整數。
    intY(8) = intX(7) Mod 10                 ' 第 7位數的乘積可能大於10，除以10，取其餘數。
  
    intSum = intX(1) + intX(3) + intX(5) + intX(8) + _
             intY(1) + intY(2) + intY(3) + intY(4) + intY(5) + intY(6) + intY(7) + intY(8)
    
    intMod = intSum Mod 10
  
    If Val(Mid(strBAN, 7, 1)) = 7 Then       ' 判斷 1: 第 7位數是否為 7 時，
       If intMod = 0 Then                    ' 判斷 2: 餘數是否為 0。
           BANCheck = True
           Exit Function
       Else
           intSum = intSum + 1               ' 再行計算。 1999/11/19 修正。
           intMod = intSum Mod 10
           If intMod = 0 Then
               BANCheck = True
               Exit Function
           Else
               strReason = "輸入錯誤，請再檢查。"
               BANCheck = False
               Exit Function
           End If
       End If
    Else
       If intMod = 0 Then
           BANCheck = True
           Exit Function
       Else
           strReason = "輸入錯誤，請再檢查。"
           BANCheck = False
           Exit Function
       End If
    End If
End Function


