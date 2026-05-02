Attribute VB_Name = "ShellExecTool"
Option Explicit
'*************************************************************************************
'專案名稱: 底層元件
'功能描述: 執行wsh shell函式
'
'版權所有: 台灣銀行
'程式撰寫: Dunk
'撰寫日期：2017/9/20
'
'改版日期:
'改版備註:
'
'*************************************************************************************

' Windows API function declarations.
#If VBA7 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, _
                                                                     ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
    Private Declare PtrSafe Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
    Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, _
                                                                             ByVal dwMilliseconds As Long) As Long
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, _
                                                                            ByRef lpExitCodeOut As Long) As Integer
#Else
    Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, _
                                                             ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, _
                                                                     ByVal dwMilliseconds As Long) As Long
    Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, _
                                                                    ByRef lpExitCodeOut As Long) As Integer
#End If

'執行wsh
Function RunWshShell(ByVal cmd As String)

    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1
    Dim errorCode As Long
    errorCode = wsh.Run(cmd, windowStyle, waitOnReturn)
    Set wsh = Nothing
    RunWshShell = errorCode
End Function

'執行wsh
Function RunWshShellHidden(ByVal cmd As String)

    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 0
    Dim errorCode As Long
    errorCode = wsh.Run(cmd, windowStyle, waitOnReturn)
    Set wsh = Nothing
    RunWshShellHidden = errorCode
End Function

' Synchronously executes the specified command and returns its exit code.
' Waits indefinitely for the command to finish, unless you pass a
' timeout value in seconds for `timeoutInSecs`.
Private Function SyncShell(ByVal cmd As String, _
                           Optional ByVal windowStyle As VbAppWinStyle = vbMinimizedFocus, _
                           Optional ByVal timeoutInSecs As Double = -1) As Long
    Dim pid As Long    ' PID (process ID) as returned by Shell().
    Dim h As Long    ' Process handle
    Dim sts As Long    ' WinAPI return value
    Dim timeoutMs As Long    ' WINAPI timeout value
    Dim exitCode As Long
    ' Invoke the command (invariably asynchronously) and store the PID returned.
    ' Note that this invocation may raise an error.
    pid = Shell(cmd, windowStyle)
    ' Translate the PIP into a process *handle* with the
    ' SYNCHRONIZE and PROCESS_QUERY_LIMITED_INFORMATION access rights,
    ' so we can wait for the process to terminate and query its exit code.
    ' &H100000 == SYNCHRONIZE, &H1000 == PROCESS_QUERY_LIMITED_INFORMATION
    h = OpenProcess(&H100000 Or &H1000, 0, pid)
    If h = 0 Then
        Err.Raise vbObjectError + 1024, , _
                  "Failed to obtain process handle for process with ID " & pid & "."
    End If
    ' Now wait for the process to terminate.
    If timeoutInSecs = -1 Then
        timeoutMs = &HFFFF    ' INFINITE
    Else
        timeoutMs = timeoutInSecs * 1000
    End If
    sts = WaitForSingleObject(h, timeoutMs)
    If sts <> 0 Then
        Err.Raise vbObjectError + 1025, , _
                  "Waiting for process with ID " & pid & _
                  " to terminate timed out, or an unexpected error occurred."
    End If
    ' Obtain the process's exit code.
    sts = GetExitCodeProcess(h, exitCode)    ' Return value is a BOOL: 1 for true, 0 for false
    If sts <> 1 Then
        Err.Raise vbObjectError + 1026, , _
                  "Failed to obtain exit code for process ID " & pid & "."
    End If
    CloseHandle h
    ' Return the exit code.
    SyncShell = exitCode
End Function
