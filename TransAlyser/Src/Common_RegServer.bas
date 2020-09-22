Attribute VB_Name = "Common_RegisterServerModule"
Option Explicit

Private Const CREATE_SUSPENDED = &H4
Private Const INFINITE = &HFFFFFFFF   ' Infinite timeout
Private Const STATUS_WAIT_0 = &H0
Private Const STATUS_ABANDONED_WAIT_0 = &H80
Private Const STATUS_TIMEOUT = &H102
Private Const WAIT_FAILED = &HFFFFFFFF
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
Private Const WAIT_ABANDONED = ((STATUS_ABANDONED_WAIT_0) + 0)
Private Const WAIT_TIMEOUT = STATUS_TIMEOUT
Private Const STATUS_PENDING = &H103
Private Const STILL_ACTIVE = STATUS_PENDING

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" _
        (ByVal lpLibFileName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
        (ByVal hLibModule As Long) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
        (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function CreateThread Lib "kernel32" _
        (lpThreadAttributes As Any, ByVal dwStackSize As Long, _
        lpStartAddress As Long, lpParameter As Any, _
        ByVal dwCreationFlags As Long, lpThreadID As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long

Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)

Private Declare Function ResumeThread Lib "kernel32" _
        (ByVal hThread As Long) As Long

Private Declare Function GetExitCodeThread Lib "kernel32" _
        (ByVal hThread As Long, lpExitCode As Long) As Long

Public Function RegServer(ByVal FilePath As String, Optional ByVal InRegister = True) As Boolean
On Error GoTo Error

    Dim lngModuleHandle As Long      ' module handle
    Dim lngFunctionAdr  As Long      ' reg/unreg function address
    Dim lngThreadID     As Long      ' dummy var that get's filled
    Dim lngThreadHandle As Long      ' thread handle
    Dim lngExitCode     As Long      ' thread's exit code if it doesn't finish
    Dim Success         As Boolean   ' if things worked

    '
    ' Load the file into memory.
    '
    lngModuleHandle = LoadLibrary(FilePath)
      
    '
    ' Get the registration function's address.
    '
    If InRegister Then
        lngFunctionAdr = GetProcAddress(lngModuleHandle, "DllRegisterServer")
    Else
        lngFunctionAdr = GetProcAddress(lngModuleHandle, "DllUnregisterServer")
    End If
    
    If lngFunctionAdr <> 0 Then
        '
        ' Create an alive thread and execute the function.
        '
        lngThreadHandle = CreateThread(ByVal 0, 0, ByVal lngFunctionAdr, ByVal 0, 0, lngThreadID)
        
        '
        ' If we got the thread handle...
        '
        If lngThreadHandle Then
            '
            ' Wait for the thread to finish.
            '
            Success = (WaitForSingleObject(lngThreadHandle, 10000) = WAIT_OBJECT_0)
          
            '
            ' If it didn't finish...
            '
            If Not Success Then
                '
                ' Something happened. Close the thread.
                '
                Call GetExitCodeThread(lngThreadHandle, lngExitCode)
                Call ExitThread(lngExitCode)
            End If
    
            '
            ' Close the thread.
            '
            Call CloseHandle(lngThreadHandle)
        End If
    End If
    
    '
    ' Free the file if we loaded it.
    '
    If lngModuleHandle Then Call FreeLibrary(lngModuleHandle)
    
    RegServer = Success

Exit Function
Error:
    Assert , "RegisterServerModule.RegServer", Err.Number, Err.Description, "FilePath: '" & FilePath & "', InRegister: '" & InRegister & "'"
    Resume Next
End Function

Public Function IsDLLActiveX(ByVal strDLLPath As String, Optional ByVal RaiseError As Boolean) As Boolean
On Error GoTo Error

    Dim lngHMod         As Long
    Dim lngLastDllError As Long
  
    lngHMod = LoadLibrary(strDLLPath)
    
    If lngHMod = 0 Then
        If RaiseError Then
            lngLastDllError = Err.LastDllError
            Err.Raise 10000 + lngLastDllError, "IsDLLActiveX", "LoadLibrary-Error: " & lngLastDllError
        End If
    End If
  
    IsDLLActiveX = Abs(CBool(GetProcAddress(lngHMod, "DllRegisterServer")))
    Call FreeLibrary(lngHMod)
    
Exit Function
Error:
    Assert , "RegisterServerModule.IsDLLActiveX", Err.Number, Err.Description, "DLLPath: '" & strDLLPath & "', RaiseError: '" & RaiseError & "'"
    Resume Next
End Function

