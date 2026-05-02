Attribute VB_Name = "CopyCliboardUtility"
Option Explicit
'*************************************************************************************
'專案名稱: 底層元件
'功能描述: 複製文字到剪貼簿
'
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2018/2/23
'
'改版日期:
'改版備註: 2020/09/17 調整複製寫法
'
'*************************************************************************************

#If Mac Then
    ' ignore
#Else
    #If VBA7 Then
        Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                                             ByVal dwBytes As LongPtr) As LongPtr

        Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
        Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
        Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long

        Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
                                                         ByVal lpString2 As Any) As LongPtr

        Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat _
                                                                As Long, ByVal hMem As LongPtr) As LongPtr
        Declare PtrSafe Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As LongPtr
    #Else
        Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
        Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
        Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
                                                     ByVal dwBytes As Long) As Long

        Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
        Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
        Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long

        Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
                                                 ByVal lpString2 As Any) As Long

        Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat _
                                                        As Long, ByVal hMem As Long) As Long
        Declare PtrSafe Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As LongPtr
    #End If
#End If
Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096

Private Sub ClipBoard_SetData(MyString As String)
    #If Mac Then
        With New MSForms.DataObject
            .SetText MyString
            .PutInClipboard
        End With
    #Else
        #If VBA7 Then
            Dim hGlobalMemory As LongPtr
            Dim hClipMemory   As LongPtr
            Dim lpGlobalMemory    As LongPtr
        #Else
            Dim hGlobalMemory As Long
            Dim hClipMemory   As Long
            Dim lpGlobalMemory    As Long
        #End If

        Dim x                 As Long

        ' Allocate moveable global memory.
        '-------------------------------------------
        hGlobalMemory = GlobalAlloc(GHND, lstrlen(MyString) + 1)

        ' Lock the block to get a far pointer
        ' to this memory.
        lpGlobalMemory = GlobalLock(hGlobalMemory)

        ' Copy the string to this global memory.
        lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

        ' Unlock the memory.
        If GlobalUnlock(hGlobalMemory) <> 0 Then
            MsgBox "Could not unlock memory location. Copy aborted."
            GoTo OutOfHere2
        End If

        ' Open the Clipboard to copy data to.
        If OpenClipboard(0&) = 0 Then
            MsgBox "Could not open the Clipboard. Copy aborted."
            Exit Sub
        End If

        ' Clear the Clipboard.
        x = EmptyClipboard()

        ' Copy the data to the Clipboard.
        hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:

        If CloseClipboard() = 0 Then
            MsgBox "Could not close Clipboard."
        End If
    #End If

End Sub

'複製到剪貼簿
'Microsoft Forms 2.0 Object Library 設定引用項目(C:\WINDOWS\system32\FM20.DLL)
Function CopyToClipboard(ByVal resultFileName As String)

'    Dim myDobj As New DataObject
'
'    '複製到剪貼簿
'    With myDobj
'        .SetText resultFileName
'        .PutInClipboard
'    End With
'    Set myDobj = Nothing
    ClipBoard_SetData resultFileName

End Function
