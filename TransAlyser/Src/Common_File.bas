Attribute VB_Name = "mFile"
Option Explicit

Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long

Private Declare Function GetShortPathName Lib "kernel32" _
      Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
      ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    If lRetVal = 0 Then 'The file does not exist, first create it!
        Open sLongFileName For Random As #1
        Close #1
        lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
        'Now another try!
        Kill (sLongFileName)
        'Delete file now!
    End If
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)
End Function

Function FileExist(FilePath As String) As Boolean
On Error GoTo Error
    If DirExist(FilePath) Then GoTo Error
    FileLen FilePath
    FileExist = True
Exit Function
Error:

End Function

Function DirExist(BaseDirPath As String) As Boolean
On Error GoTo Error
    If InStr(1, BaseDirPath, "\\") <> 0 Then GoTo Error
    If (GetAttr(BaseDirPath) And vbDirectory) = vbDirectory Then DirExist = True
Exit Function
Error:
     DirExist = False
End Function

Function GetFileList(Path As String, FileArray() As String, Optional WithPath As Boolean = True, Optional IncludeSubFolders As Boolean) As Long
Dim FileName As String, sPath As String
Dim Files As Long
On Error GoTo Error
    sPath = FixPath(Path)
    FileName = Dir(sPath, vbArchive + vbHidden + vbReadOnly + vbSystem)
    Do
        Select Case FileName
            Case "", ".", ".."
            Case Else
                Files = Files + 1
                ReDim Preserve FileArray(Files)
                If WithPath Then
                    FileArray(Files) = sPath & FileName
                Else
                    FileArray(Files) = FileName
                End If
        End Select
        On Error Resume Next
        FileName = Dir()
        On Error GoTo Error
    Loop While FileName <> ""
    If Files = 0 Then ReDim FileArray(0)
    GetFileList = Files
Exit Function
Error:
    Assert , "FileModule.GetFileList", Err.Number, Err.Description, "Path: '" & Path & "'"
    Resume Next
End Function

Function GetDirList(Path As String, DirArray() As String) As Long
Dim DirName As String, BaseDirPath As String, p As String
Dim Dirs As Long
On Error GoTo Error
    If DirExist(Path) = False Then GoTo Error
    If Right(Path, 1) <> "\" Then
        BaseDirPath = Path & "\"
    Else
        BaseDirPath = Path
    End If
    DirName = Dir(BaseDirPath, vbDirectory + vbArchive + vbHidden + vbReadOnly + vbSystem)
    Do
        DoEvents
        Select Case DirName
            Case "", ".", ".."
            Case Else
                p = BaseDirPath & DirName
                If (GetAttr(p) And vbDirectory) = vbDirectory Then
                    Dirs = Dirs + 1
                    ReDim Preserve DirArray(Dirs)
                    DirArray(Dirs) = DirName
                End If
        End Select
        DirName = Dir()
    Loop While DirName <> ""
    GetDirList = Dirs
    
Exit Function
Error:
    Assert , "FileModule.GetDirList", Err.Number, Err.Description, "Path: '" & Path & "'"
    Resume Next
End Function

Function GetSubDirList(Path As String, DirArray() As String) As Long
Dim i As Long, x As Long
Dim sPath As String
Dim DirCount&
Dim SubDirCount1&
Dim SubDirCount2&
Dim SubDirList1() As String
Dim SubDirList2() As String
On Error GoTo Error
    SubDirCount1 = GetDirList(Path, SubDirList1())
    For i = 1 To SubDirCount1
        DoEvents
        sPath = FixPath(Path) & SubDirList1(i)
        DirCount = DirCount + 1
        ReDim Preserve DirArray(DirCount)
        DirArray(DirCount) = sPath
        SubDirCount2 = GetSubDirList(sPath, SubDirList2())
        For x = 1 To SubDirCount2
            DoEvents
            DirCount = DirCount + 1
            ReDim Preserve DirArray(DirCount)
            DirArray(DirCount) = SubDirList2(x)
        Next x
    Next i
    GetSubDirList = DirCount
Exit Function
Error:
    Assert , "FileModule.GetSubDirList", Err.Number, Err.Description, "Path: '" & Path & "'"
    Resume Next
End Function

Function DeleteFile(FilePath As String) As Boolean
On Error GoTo Error
    SetAttr FilePath, vbNormal
    Kill FilePath
    DeleteFile = True
Exit Function
Error:
    Assert , "FileModule.DeleteFile", Err.Number, Err.Description, "FilePath: '" & FilePath & "'", False
End Function

Function DeleteDir(sFolder As String) As Boolean
Dim sCurrFile As String, sFilePath As String
Dim lAttribs As Long
On Error GoTo Error
    lAttribs = vbDirectory + vbArchive + vbHidden + vbReadOnly + vbSystem
    sCurrFile = Dir(sFolder & "\", lAttribs)
    Do While Len(sCurrFile) > 0
        If sCurrFile <> "." And sCurrFile <> ".." Then
            If (GetAttr(sFolder & "\" & sCurrFile) And vbDirectory) = vbDirectory Then
                DeleteDir sFolder & "\" & sCurrFile
                sCurrFile = Dir(sFolder & "\", lAttribs)
            Else
                sFilePath = sFolder & "\" & sCurrFile
                SetAttr sFilePath, vbNormal
                Kill sFilePath
                sCurrFile = Dir
            End If
        Else
            sCurrFile = Dir
        End If
        DoEvents
    Loop
    SetAttr sFolder, vbNormal
    RmDir sFolder
    DeleteDir = True
Exit Function
Error:
    
End Function

Function FreeDiskSpace(sDrive As String) As Double
On Error GoTo Error
    
    Dim strRootPathName
    Dim lngSectorsPerCluster
    Dim lngBytesPerSector
    Dim lngNumberOfFreeClusters
    Dim lngTotalNumberOfClusters
    Dim strDrive
    Dim strMessage
    Dim lngTotalBytes
    Dim lngFreeBytes

    strDrive = Left(sDrive, 2)
    GetDiskFreeSpace strDrive, lngSectorsPerCluster, lngBytesPerSector, lngNumberOfFreeClusters, lngTotalNumberOfClusters
    lngFreeBytes = lngNumberOfFreeClusters * lngSectorsPerCluster * lngBytesPerSector
    FreeDiskSpace = lngFreeBytes

Exit Function
Error:
    Assert , "FileModule.FreeDiskSpace", Err.Number, Err.Description, "Drive: '" & sDrive & "'"
    Resume Next
End Function

Function IsFilePath(sPath As String) As Boolean
On Error GoTo Error
    If FileExist(sPath) Then
        IsFilePath = True
        Exit Function
    End If
    WriteFile sPath, sPath, , , True
    If Not FileExist(sPath) Then GoTo Error
    DeleteFile sPath
    IsFilePath = True
    Exit Function
Error:

End Function

Function ReadFile(FilePath As String, Optional Start As Long, Optional Lenght As Long, Optional Reverse As Boolean = False) As String
On Error GoTo Error

    Dim FileNum As Long
    Dim Buffer As String
    Dim lLen As Long
    Dim lStart As Long
    Dim lLOF As Long

    lLen = Lenght
    lStart = Start
    If Not FileExist(FilePath) Then
        Err.Description = "File doesn't exist"
        GoTo Error
    End If
    FileNum = FreeFile
    Open FilePath For Binary As #FileNum
        lLOF = LOF(FileNum)
        If lStart > lLOF Then
            Err.Description = "Startpoint > Filelenght"
            GoTo Error
        End If
        If lStart = 0 Then lStart = 1
        If lLen = 0 Or (lStart + lLen - 1) > lLOF Then lLen = (lLOF - lStart + 1)
        If Reverse Then lStart = lLOF - lStart - lLen + 2
        Buffer = String(lLen, " ")
        Get #FileNum, lStart, Buffer
    Close FileNum
    ReadFile = Buffer
    
Exit Function
Error:
    Assert , "FileModule.ReadFile", Err.Number, Err.Description, "FilePath: '" & FilePath & "', Start: '" & CStr(Start) & "', Lenght: '" & Lenght & "', Reverse: '" & CStr(Reverse) & "'"
    Resume Next
End Function

Function WriteFile(sFilePath$, sData$, Optional ByVal lStart&, Optional Insert As Boolean, Optional NoError As Boolean)
On Error GoTo Error

    Dim FileNum As Long
    Dim sBuffer As String
    Dim FileExists As Boolean
    Dim FileSize As Long
    Dim Buffer As String
    
    If Not DirExist(GetDirName(sFilePath)) Then
        Assert "Can not write file because dir """ & GetDirName(sFilePath) & """ does not exist."
        Exit Function
    End If
    
    If NoError Then On Error Resume Next
    
    FileExists = FileExist(sFilePath)
    If FileExists Then
        FileSize = FileLen(sFilePath)
    End If
    If lStart < 1 Then
        If FileExists Then lStart = FileSize + 1
        If lStart < 1 Then lStart = 1
    End If
    FileNum = FreeFile
    If FileExists And Insert Then
        Buffer = sData & ReadFile(sFilePath, lStart)
        Open sFilePath For Binary As FileNum
        Put FileNum, lStart, Buffer
        Close FileNum
    Else
        Open sFilePath For Binary As FileNum
        Put FileNum, lStart, sData
        Close FileNum
    End If
        
Exit Function
Error:
    Assert , "FileModule.WriteFile", Err.Number, Err.Description, "FilePath: '" & sFilePath & "', Start: '" & CStr(lStart) & "', Insert: '" & CStr(Insert) & "'"
    Resume Next
End Function

Function GetFileName(FilePath As String, Optional Delimiter As String = "\") As String
On Error GoTo Error

    Dim i As Long
    
    If InStr(1, FilePath, Delimiter) = 0 Then
        GetFileName = FilePath
        Exit Function
    End If
    
    For i = 1 To Len(FilePath)
        If Left(Right(FilePath, i), 1) = Delimiter Then
            GetFileName = Right(FilePath, i - 1)
            Exit Function
        End If
    Next i
    
    GetFileName = FilePath

Exit Function
Error:
    Assert , "FileModule.GetFileName", Err.Number, Err.Description, "FilePath: '" & FilePath & "', Delimiter: '" & Delimiter & "'"
    Resume Next
End Function

Function GetFileTitle(FileName As String, Optional ExtLen As Integer = 3)
    GetFileTitle = GetFileName(Left(FileName, Len(FileName) - (ExtLen + 1)))
End Function

Function GetDirName(FilePath As String) As String
On Error GoTo Error

    GetDirName = Left(FilePath, Len(FilePath) - Len(GetFileName(FilePath)))

Exit Function
Error:
    Assert , "FileModule.GetDirName", Err.Number, Err.Description, "FilePath: '" & FilePath & "'"
    Resume Next
End Function

Function BuildPath(Path As String, Optional Test As Boolean) As Boolean
Dim DirArray() As String, BuiltDirArray() As String, cDir As String
Dim d As Integer, i As Integer
On Error GoTo Error
    DirArray() = Split(FixPath(Path), "\")
    For i = 0 To UBound(DirArray())
        If i <> 0 Then cDir = cDir & "\"
        cDir = cDir & DirArray(i)
        If Not DirExist(cDir) Then
            MkDir cDir
            ReDim Preserve BuiltDirArray(d)
            BuiltDirArray(d) = cDir
            d = d + 1
        End If
    Next i
    BuildPath = True
    If Not Test Then Exit Function
Error:
On Error Resume Next
    If d > 0 Then
        For i = 1 To d
            RmDir BuiltDirArray(d - i)
        Next i
    End If
    If Not Test Then
        Assert , "FileModule.BuildPath", Err.Number, Err.Description, "Path: '" & Path & "', Test: '" & CStr(Test) & "'"
        Resume Next
    End If
End Function

Public Function SaveFileAttachment(FilePath As String, sArray() As String)
On Error GoTo Error
    
    WriteFile FilePath, JoinArrayRev(sArray), FileLen(FilePath)
    
Exit Function
Error:
    Assert , "FileModule.SaveFileAttachment", Err.Number, Err.Description, "FilePath: '" & FilePath & "'"
    Resume Next
End Function

Public Function LoadFileAttachment(sFilePath$, sArray$())
On Error GoTo Error
    
    Dim sTOC$
    Dim sRecordLen$()
    Dim lRecordLen&()
    Dim lAttachLen&
    Dim sAttach$
    Dim i&
    
    Do
        i = i + 1
        sTOC = ReadFile(sFilePath, i, 1, True) & sTOC
    Loop While Left(sTOC, 1) <> Chr(255)
    sRecordLen() = Split(Right(sTOC, Len(sTOC) - 1), Chr(254))
    ReDim lRecordLen(LBound(sRecordLen) To UBound(sRecordLen))
    For i = LBound(sRecordLen) To UBound(sRecordLen)
        lRecordLen(i) = StringToNumber(sRecordLen(i))
    Next i
    lAttachLen = lSumOfArray(lRecordLen) + Len(sTOC)
    sAttach = ReadFile(sFilePath, 1, lAttachLen, True)
    SplitArrayRev sAttach, sArray()

Exit Function
Error:
    Assert , "FileModule.LoadFileAttachment", Err.Number, Err.Description, "FilePath: '" & sFilePath & "'"
    Resume Next
End Function

Function CountSectors(Point As Long, SectorSize As Long) As Long
On Error GoTo Error

    CountSectors = Point \ SectorSize
    If (Point Mod SectorSize) > 0 Then CountSectors = CountSectors + 1

Exit Function
Error:
    Assert , "FileModule.CountSectors", Err.Number, Err.Description, "Point: '" & CStr(Point) & "', SectorSize: '" & CStr(SectorSize) & "'"
    Resume Next
End Function

Function SectorStart(Sector As Long, SectorSize As Long)
On Error GoTo Error

    SectorStart = ((Sector - 1) * SectorSize) + 1

Exit Function
Error:
    Assert , "FileModule.SectorStart", Err.Number, Err.Description, "Sector: '" & CStr(Sector) & "', SectorSize: '" & CStr(SectorSize) & "'"
    Resume Next
End Function

Function CalcSectorSize(MaxPoint As Long, Sector As Long, SectorSize As Long)
On Error GoTo Error
    
    If CountSectors(MaxPoint, SectorSize) <> Sector Then
        CalcSectorSize = SectorSize
    Else
        CalcSectorSize = SectorSize - ((Sector * SectorSize) - MaxPoint)
    End If
    
Exit Function
Error:
    Assert , "FileModule.CalcSectorSize", Err.Number, Err.Description, "MaxPoint: '" & CStr(MaxPoint) & "', Sector: '" & CStr(Sector) & "', SectorSize: '" & CStr(SectorSize) & "'"
    Resume Next
End Function

Function FormatBytes(NumBytes As Double, Optional Magnitude As String) As String
    If Magnitude = "" Then
        Select Case NumBytes
            Case 0 To 1024
                FormatBytes = CStr(NumBytes) & " Bytes"
            Case (1024 ^ 1) + 1 To (1024 ^ 2)
                FormatBytes = Format(CStr(NumBytes / 1024), "fixed") & " KB"
            Case (1024 ^ 2) + 1 To (1024 ^ 3)
                FormatBytes = Format(CStr(NumBytes / 1024 ^ 2), "fixed") & " MB"
            Case (1024 ^ 3) + 1 To (1024 ^ 4)
                FormatBytes = Format(CStr(NumBytes / 1024 ^ 3), "fixed") & " GB"
            Case (1024 ^ 4) + 1 To (1024 ^ 5)
                FormatBytes = Format(CStr(NumBytes / 1024 ^ 4), "fixed") & " TB"
        End Select
    Else
        Select Case Magnitude
            Case "Bytes"
                FormatBytes = CStr(NumBytes) & " Bytes"
            Case "KB"
                FormatBytes = Format(CStr(NumBytes / 1024), "fixed") & " KB"
            Case "MB"
                FormatBytes = Format(CStr(NumBytes / 1024 ^ 2), "fixed") & " MB"
            Case "GB"
                FormatBytes = Format(CStr(NumBytes / 1024 ^ 3), "fixed") & " GB"
            Case "TB"
                FormatBytes = Format(CStr(NumBytes / 1024 ^ 4), "fixed") & " TB"
        End Select
    End If
End Function
