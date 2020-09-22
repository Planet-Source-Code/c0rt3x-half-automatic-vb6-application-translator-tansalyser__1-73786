Attribute VB_Name = "mProcess"
Option Explicit

Private Declare Function GetClassLong Lib "user32" _
 Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) _
 As Long
Private Declare Function GetModuleFileName Lib "kernel32" _
 Alias "GetModuleFileNameA" (ByVal hModule As Long, _
 ByVal lpFileName As String, ByVal nSize As Long) As Long
  
Private Const GCL_HMODULE = (-16)

Private Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" _
   (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpLayout As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, _
   lpProcessInformation As PROCESS_INFORMATION) As Long

Public Enum eProcessPriorityClasses
    [ePPC - Idle] = 64
    [ePPC - Normal] = 32
    [ePPC - High] = 128
    [ePPC - Realtime] = 256
End Enum

Private Const STARTF_USESHOWWINDOW = &H1

Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000

Private Declare Function CloseHandle Lib "kernel32" (hObject As Long) As Boolean
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadID As Long
End Type

Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const MAX_PATH As Integer = 260

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" _
    Alias "CreateToolhelp32Snapshot" _
   (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Private Declare Function ProcessFirst Lib "kernel32" _
    Alias "Process32First" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Function ProcessNext Lib "kernel32" _
    Alias "Process32Next" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Public Function GetAppEXEPath() As String
On Error GoTo Error
    GetAppEXEPath = GetModulePath
Exit Function
Error:
    Assert , "ProcessModule.GetAppExePath", Err.Number, Err.Description
End Function

Public Function GetModulePath(Optional hwnd As Long) As String
On Error GoTo Error
    
    Dim nResult As String
    Dim nLenResult As Long
    
    nResult = Space(260)
    nLenResult = Len(nResult)
    nLenResult = GetModuleFileName(GetClassLong(hwnd, GCL_HMODULE), _
    nResult, nLenResult)
    GetModulePath = Left$(nResult, nLenResult)
    
Exit Function
Error:
    Assert , "ProcessModule.GetModulePath", Err.Number, Err.Description
    Resume Next
End Function

Public Function HideProc(HandleForm As Form)
On Error GoTo Error

    Dim lHandle&
    
    RegisterServiceProcess GetCurrentProcessId, 1
    lHandle = GetWindow(HandleForm.hwnd, 4)
    ShowWindow lHandle, 0
    
Exit Function
Error:
    Assert , "ProcessModule.HideProc", Err.Number, Err.Description, "HandleForm: '" & HandleForm.Name & "'"
    Resume Next
End Function

Public Function Run(sCommand As String, Optional Sync As Boolean, Optional WindowStyle As VbAppWinStyle = vbNormalFocus, Optional Priority As eProcessPriorityClasses = [ePPC - Normal])
On Error GoTo Error

    Dim ProcInfo As PROCESS_INFORMATION
    Dim StartInfo As STARTUPINFO
    Dim i As Long
    Dim Ret As Long
    
    With StartInfo
        .dwFlags = STARTF_USESHOWWINDOW
        .wShowWindow = WindowStyle
        .cb = Len(StartInfo)
    End With
    Ret = CreateProcessA(0&, sCommand, 0&, 0&, 1&, Priority, 0&, 0&, StartInfo, ProcInfo)
    With ProcInfo
        If Ret <> 0 Then
            If Sync Then
                Ret = WaitForSingleObject(.hProcess, INFINITE)
                GetExitCodeProcess .hProcess, Ret
                CloseHandle .hThread
                CloseHandle .hProcess
            End If
        End If
    End With
    Run = Ret

Exit Function
Error:
    Assert , "ProcessModule.Run", Err.Number, Err.Description, "Command: '" & sCommand & "', Sync: '" & CStr(Sync) & "'"
    Resume Next
End Function

Public Function GetProcessInstances(ProcPath As String, hInstance() As Long)
On Error GoTo Error

    Dim hSnapShot&
    Dim Proc As PROCESSENTRY32
    Dim r&, i&
    Dim ProcessID&, hProcess&
    
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = 0 Then Exit Function
    
    Proc.dwSize = Len(Proc)
    
    r = ProcessFirst(hSnapShot, Proc)
    
    ReDim hInstance(0)
    
    Do While r
        
        r = ProcessNext(hSnapShot, Proc)
        
        If InStr(1, LCase(Proc.szExeFile), LCase(GetFileName(ProcPath))) <> 0 Then
            i = i + 1
            ReDim Preserve hInstance(i)
            hInstance(i) = Proc.th32ProcessID
        End If
        
    Loop
    
    CloseHandle hSnapShot

Exit Function
Error:
    Assert , "ProcessModule.GetProcessInstances", Err.Number, Err.Description, "ProcPath: '" & ProcPath & "'"
    Resume Next
End Function

Public Function Suicide(Optional AllInstances As Boolean)
On Error GoTo Error

    If AllInstances Then
        KillAllInstances GetAppEXEPath
    Else
        KillProc GetCurrentProcessId
    End If
    End

Exit Function
Error:
    Assert , "ProcessModule.Suicide", Err.Number, Err.Description, "AllInstances: '" & CStr(AllInstances) & "'"
    Resume Next
End Function

Public Function KillAllInstances(ProcPath As String)
On Error GoTo Error

    Dim i&, c&
    Dim ThisInstance As Long
    Dim hInstance() As Long
    Dim hProcess&
    Dim ExitCode As Long
    
    ThisInstance = GetCurrentProcessId
            
    GetProcessInstances ProcPath, hInstance()
    
    For i = 1 To UBound(hInstance)
        If hInstance(i) <> ThisInstance Then
            KillProc hInstance(i)
            c = c + 1
        End If
    Next i
    
    If c = UBound(hInstance()) Then Exit Function
    
    KillProc ThisInstance

Exit Function
Error:
    Assert , "ProcessModule.KillAllInstances", Err.Number, Err.Description, "ProcPath: '" & ProcPath & "'"
    Resume Next
End Function

Public Function KillProc(ProcID As Long)
On Error GoTo Error

    Dim hInstance As Long
    Dim hProcess As Long
    Dim ExitCode As Long
    
    hProcess = OpenProcess(&H1, 1, ProcID)
    GetExitCodeProcess hProcess, ExitCode
    TerminateProcess hProcess, ExitCode

Exit Function
Error:
    Assert , "ProcessModule.KillProc", Err.Number, Err.Description
    Resume Next
End Function
