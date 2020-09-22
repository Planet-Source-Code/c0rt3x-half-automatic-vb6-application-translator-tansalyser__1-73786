Attribute VB_Name = "Common_SystemModule"
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal lLocale As Long, ByVal lLocaleType As Long, ByVal sLCData As String, ByVal lBufferLength As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WM_ACTIVATEAPP = &H1C
Private Const GWL_WNDPROC = -4

Private lOldWndProc As Long
Private mForm As Object

Private Const BITSPIXEL = 12&
Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000
Private Const CDS_TEMP = &H4

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
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
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

'Private Declare Function GetDesktopWindow Lib "User32" () As Long
'Private Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long

Function GetSystemLanguage() As String
On Error GoTo Error
    
    Dim lBuffSize As String, sBuffer As String
    Dim lRet As Long

    lBuffSize = 256
    sBuffer = String$(lBuffSize, vbNullChar)
    lRet = GetLocaleInfo(1024, 4, sBuffer, lBuffSize)
    If lRet > 0 Then
        GetSystemLanguage = Left$(sBuffer, lRet - 1)
    End If
    
Exit Function
Error:
    Assert , "SystemModule.GetSystemLanguage", Err.Number, Err.Description
    Resume Next
End Function

Function GetSystemCountry() As String
On Error GoTo Error
    
    Dim lBuffSize As String, sBuffer As String
    Dim lRet As Long
    
    lBuffSize = 256
    sBuffer = String$(lBuffSize, vbNullChar)
    lRet = GetLocaleInfo(1024, 6, sBuffer, lBuffSize)
    If lRet > 0 Then
        GetSystemCountry = Left$(sBuffer, lRet - 1)
    End If
    
Exit Function
Error:
    Assert , "SystemModule.GetSystemCountry", Err.Number, Err.Description
    Resume Next
End Function

Public Function GetScreenResolution(PixelX As Long, PixelY As Long)
On Error GoTo Error
 
    Dim DeskRect As RECT
 
    GetWindowRect GetDesktopWindow, DeskRect
    PixelX = DeskRect.Right
    PixelY = DeskRect.Bottom
    
Exit Function
Error:
    Assert , "SystemModule.GetScreenResolution", Err.Number, Err.Description
    Resume Next
End Function

Public Function IsWin9x() As Boolean
On Error GoTo Error

    Dim Info As OSVERSIONINFO
    Info.dwOSVersionInfoSize = Len(Info)
    GetVersionEx Info
    IsWin9x = (Info.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS)

Exit Function
Error:
    Assert , "SystemModule.IsWin9x", Err.Number, Err.Description
    Resume Next
End Function

Public Function GetScreenRes(WidthPixels As Long, HeightPixels As Long, BitDepth As Long)
On Error GoTo Error
    
    Dim vR As RECT
    
    Call GetWindowRect(GetDesktopWindow, vR)
    WidthPixels = vR.Right
    HeightPixels = vR.Bottom
    BitDepth = GetDeviceCaps(GetDC(GetDesktopWindow), BITSPIXEL)

Exit Function
Error:
    Assert , "SystemModule.GetScreenRes", Err.Number, Err.Description
    Resume Next
End Function

Public Function SetScreenRes(ByVal WidthPixels As Long, ByVal HeightPixels As Long, ByVal BitDepth As Long) As Boolean
On Error GoTo Error
    
    Dim DevM As DEVMODE
    
    EnumDisplaySettings 0&, 0&, DevM
    
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = WidthPixels
    DevM.dmPelsHeight = HeightPixels
    DevM.dmBitsPerPel = BitDepth
    SetScreenRes = (ChangeDisplaySettings(DevM, CDS_TEMP) = 0)

Exit Function
Error:
    Assert , "SystemModule.SetScreenRes", Err.Number, Err.Description, "WidthPixels: '" & WidthPixels & "', HeightPixels: '" & HeightPixels & "', BitDepth: '" & BitDepth & "'"
    Resume Next
End Function

Public Function WinVersion() As String
    
    Dim udtOSVersion As OSVERSIONINFO
    Dim lMajorVersion  As Long
    Dim lMinorVersion As Long
    Dim lPlatformID As Long
    Dim sAns As String
    
    
    udtOSVersion.dwOSVersionInfoSize = Len(udtOSVersion)
    GetVersionEx udtOSVersion
    lMajorVersion = udtOSVersion.dwMajorVersion
    lMinorVersion = udtOSVersion.dwMinorVersion
    lPlatformID = udtOSVersion.dwPlatformId
    
    Select Case lMajorVersion
        Case 5
        
            ' Added the following to give suppport for Windows XP!
            If lMinorVersion = 0 Then
            
                sAns = "Windows 2000"
                
            ElseIf lMinorVersion = 1 Then
            
                sAns = "Windows XP"
            
            End If
              
        Case 4
            If lPlatformID = VER_PLATFORM_WIN32_NT Then
                sAns = "Windows NT 4.0"
            Else
                sAns = IIf(lMinorVersion = 0, _
                "Windows 95", "Windows 98")
            End If
        Case 3
            If lPlatformID = VER_PLATFORM_WIN32_NT Then
                sAns = "Windows NT 3.x"
 
              'below should only happen if person has Win32s
                'installed
            Else
                sAns = "Windows 3.x"
            End If
            
        Case Else
            sAns = "Unknown Windows Version"
    End Select
                    
    WinVersion = sAns
    
End Function


