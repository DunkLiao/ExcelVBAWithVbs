Attribute VB_Name = "StringUtility"
''字串處理模組
''作者：Guan Jhih Liao

'顯示今日日期(民國年YYYMMDD)
Public Function Today() As String
  Today = Format(Date, "yyyymmdd")
  Today = Today - 19110000
  TodayTw = Format(TodayTw, "0######")
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

