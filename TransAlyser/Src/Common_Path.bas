Attribute VB_Name = "mPath"
Private Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const MAX_PATH_LEN = 260

Private Type SHITEMID
    SHItem As Long
    itemID() As Byte
End Type
Private Type ITEMIDLIST
    shellID As SHITEMID
End Type

Private Const DESKTOP = &H0
Private Const PROGRAMS = &H2
Private Const MYDOCS = &H5
Private Const FAVORITES = &H6
Private Const STARTUP = &H7
Private Const RECENT = &H8
Private Const SENDTO = &H9
Private Const STARTMENU = &HB
Private Const NETHOOD = &H13
Private Const FONTS = &H14
Private Const SHELLNEW = &H15
Private Const TEMPINETFILES = &H20
Private Const COOKIES = &H21
Private Const HISTORY = &H22

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwnd As Long, ByVal folderid As Long, shidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal shidl As Long, ByVal shPath As String) As Long

Function FixPath(Path As String, Optional Terminate As Boolean = True) As String
On Error GoTo Error

    If Right(Path, 1) = "\" Then
        FixPath = Left(Path, Len(Path) - 1)
    Else
        FixPath = Path
    End If
    If Terminate Then FixPath = FixPath & "\"

Exit Function
Error:
    Assert , "PathModule.FixPath", Err.Number, Err.Description, "Path: '" & Path & "', Terminate: '" & CStr(Terminate) & "'"
    Resume Next
End Function

Function AdaptFileName(FileName As String) As String
On Error GoTo Error
    
    Const ValidChars = "!""#$%&'()+,-0123456789;<=>@ABCDEFGHIJKLMNOPQRSTUVWXYZ[]^_`abcdefghijklmnopqrstuvwxyz{}~ÄÅÇÉÑÖÜáàâäãåçéèêëíìîïñóòôöõúùûü†°¢£§•¶ß®©™´¨≠ÆØ∞±≤≥¥µ∂∑∏π∫ªºΩæø¿¡¬√ƒ≈∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ‡·‚„‰ÂÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ"
    
    Dim i As Long, c As String * 1
    
    For i = 1 To Len(FileName)
        c = Mid(FileName, i, 1)
        If InStr(1, ValidChars, c) = 0 Then
            AdaptFileName = AdaptFileName & "_"
        Else
            AdaptFileName = AdaptFileName & c
        End If
    Next i

Exit Function
Error:
    Assert , "PathModule.AdaptFilePath", Err.Number, Err.Description, "Path: '" & Path & "', Terminate: '" & CStr(Terminate) & "'"
    Resume Next
End Function

Function GetWinPath() As String
On Error GoTo Error

    Dim s As String, i As Integer
    
    s = Space(MAX_PATH_LEN)
    i = GetWindowsDirectoryA(s, MAX_PATH_LEN)
    GetWinPath = FixPath(Left(s, i))

Exit Function
Error:
    Assert , "PathModule.GetWinPath", Err.Number, Err.Description
    Resume Next
End Function

Function GetSysPath() As String
On Error GoTo Error

    Dim s$, i%
    
    s = Space(MAX_PATH_LEN)
    i = GetSystemDirectoryA(s, MAX_PATH_LEN)
    GetSysPath = FixPath(Left(s, i))

Exit Function
Error:
    Assert , "PathModule.GetSysPath", Err.Number, Err.Description
    Resume Next
End Function

Public Function GetTempPath() As String
On Error GoTo Error

    Dim s As String, i As Integer

    s = Space(MAX_PATH_LEN)
    i = GetTempPathA(MAX_PATH_LEN, s)
    GetTempPath = FixPath(Left$(s, i))
   
Exit Function
Error:
    Assert , "PathModule.GetTempPath", Err.Number, Err.Description
    Resume Next
End Function

Function TempFilePath(Optional Extension As String = ".tmp", Optional Lenght As Integer = 8) As String
On Error GoTo Error

    Dim TempPath$
    Dim FilePath$
    
    TempPath = GetTempPath
    Do
        FilePath = TempPath & RandomFileName(Extension, Lenght)
    Loop While FileExist(FilePath)
    TempFilePath = FilePath
    
Exit Function
Error:
    Assert , "PathModule.TempFilePath", Err.Number, Err.Description, "Extension: '" & Extension & "', Lenght: '" & CStr(Lenght) & "'"
    Resume Next
End Function

Public Function GetDesktopPath() As String
On Error GoTo Error

    Dim Path As String * 256
    Dim myid As ITEMIDLIST
    Dim rval As Long
    'Get desktop path
    rval = SHGetSpecialFolderLocation(0, DESKTOP, myid)
    If rval = 0 Then ' If success
        rval = SHGetPathFromIDList(ByVal myid.shellID.SHItem, ByVal Path)
        If rval Then ' If True
            GetDesktopPath = FixPath(Left(Path, InStr(Path, Chr(0)) - 1))
        End If
    End If
   
Exit Function
Error:
    Assert , "PathModule.GetTempPath", Err.Number, Err.Description
    Resume Next
End Function

Public Function GetLongFilename(ByVal sShortName As String) As String

     Dim sLongName As String
     Dim sTemp As String
     Dim iSlashPos As Integer

     'Add \ to short name to prevent Instr from failing
     sShortName = sShortName & "\"

     'Start from 4 to ignore the "[Drive Letter]:\" characters
     iSlashPos = InStr(4, sShortName, "\")

     'Pull out each string between \ character for conversion
     While iSlashPos
       sTemp = Dir(Left$(sShortName, iSlashPos - 1), _
         vbNormal + vbHidden + vbSystem + vbDirectory)
       If sTemp = "" Then
         'Error 52 - Bad File Name or Number
         GetLongFilename = ""
         Exit Function
       End If
       sLongName = sLongName & "\" & sTemp
       iSlashPos = InStr(iSlashPos + 1, sShortName, "\")
     Wend

     'Prefix with the drive letter
     GetLongFilename = Left$(sShortName, 2) & sLongName

   End Function

