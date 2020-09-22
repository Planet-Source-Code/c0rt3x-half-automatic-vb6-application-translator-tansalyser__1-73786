VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long





Private pID                 As Long
Private pAlias              As VBAlias
Private pFilePath           As String
Private pFileAlias          As VBAlias
Private pCode               As String

Private pResFilePath        As String
Private pResFileAlias       As VBAlias

Private pLineCount          As Long
Private pLines()            As String

Private pBlockCount         As Long
Private pBlocks()           As VBCodeBlock

Private pModuleCount        As Long
Private pModules()          As VBModule

Private pStringCount        As Long
Private pStrings()          As VBString

Private pSpecialModuleCount As Long
Private pSpecialModules()   As VBSpecialModule

'Private pSpecialModuleID()  As Long

Private pExeName            As String
Private pOutDir             As String

Public Libs                 As Collection
Public LibIDs               As Collection

'Public StringEncryption     As VBStringEncryption

Private WithEvents CurrentModule As VBModule
Attribute CurrentModule.VB_VarHelpID = -1

Public Cancelled            As Boolean

Public Event FileProgress(Progress As Double)
Public Event ProjectProgress(Progress As Double)


Private Sub Class_Initialize()
    
    Set pAlias = New VBAlias
    Set pFileAlias = New VBAlias
    Set pResFileAlias = New VBAlias
    
    Set Libs = New Collection
    Set LibIDs = New Collection
    
    'Set StringEncryption = New VBStringEncryption
    
    
    With pAlias
        .AliasType = VBProject
        .Scope = VBLocal + VBPrivate + VBPublic
    End With
    
    With pFileAlias
        .AliasType = vbFile
        .Scope = VBLocal + VBPrivate + VBPublic
    End With
    
    With pResFileAlias
        .AliasType = vbFile
        .Scope = VBLocal + VBPrivate + VBPublic
    End With
    
    pBlockCount = 0
    ReDim pBlocks(0)
    
    pModuleCount = 0
    ReDim pModules(0)
    
    pStringCount = 0
    ReDim pStrings(0)
    
    pSpecialModuleCount = 0
    ReDim pSpecialModules(0)
    
End Sub

Property Get ID() As Long
    ID = pID
End Property

Property Let ID(NewValue As Long)
    pID = NewValue
    pAlias.ModuleID = pID
End Property


Property Get Name() As String
    Name = pAlias.Name
End Property

Property Let Name(NewValue As String)
    pAlias.Name = NewValue
End Property


Property Get Alias() As VBAlias
    Set Alias = pAlias
End Property


Property Get FilePath() As String
    FilePath = pFilePath
End Property

Property Let FilePath(NewValue As String)
    pFilePath = NewValue
    pFileAlias.Name = GetFileName(NewValue)
End Property

Property Get FileName() As String
    FileName = pFileAlias.Name
End Property

Property Get FileAlias() As VBAlias
    Set FileAlias = pFileAlias
End Property

Property Get LineCount() As Long
    LineCount = pLineCount
End Property

Property Get Lines(Index As Long) As String
    Lines = pLines(Index)
End Property

Property Let Lines(Index As Long, NewValue As String)
    pLines(Index) = NewValue
End Property

Property Get Code() As String
    Code = pCode
End Property

Property Let Code(NewValue As String)
    pCode = NewValue
End Property


Property Get BaseDirPath() As String
    BaseDirPath = GetDirName(FilePath)
End Property


Property Get FileTitle() As String
    FileTitle = GetFileTitle(FilePath)
End Property


Property Get ResFilePath() As String
    ResFilePath = pResFilePath
End Property

Property Let ResFilePath(NewValue As String)
    pResFilePath = NewValue
End Property



Property Get ResFileName() As String
    ResFileName = GetFileName(ResFilePath)
End Property


Property Let ResFileName(NewValue As String)
    pResFileAlias.Name = NewValue
End Property

Property Get ResFileAlias() As VBAlias
    Set ResFileAlias = pResFileAlias
End Property


Property Get BlockCount() As Long
    BlockCount = pBlockCount
End Property

Property Let BlockCount(NewValue As Long)
    ReDim Preserve pBlocks(NewValue)
    pBlockCount = NewValue
End Property


Property Get Blocks(Index As Long) As VBCodeBlock
    Set Blocks = pBlocks(Index)
End Property

Property Set Blocks(Index As Long, NewObject As VBCodeBlock)
    Set pBlocks(Index) = NewObject
End Property


Property Get ModuleCount() As Long
    ModuleCount = pModuleCount
End Property

Property Let ModuleCount(NewValue As Long)
    ReDim Preserve pModules(NewValue)
    pModuleCount = NewValue
End Property


Property Get Modules(Index As Long) As VBModule
    Set Modules = pModules(Index)
End Property

Property Set Modules(Index As Long, NewObject As VBModule)
    Set pModules(Index) = NewObject
End Property


Property Get ModuleTypeCount(DataType As String) As Long
    Dim i As Long
    For i = 1 To ModuleCount
        If Modules(i).DataType = DataType Then
            ModuleTypeCount = ModuleTypeCount + 1
        End If
    Next i
End Property


Property Get SpecialModuleCount() As Long
    SpecialModuleCount = pSpecialModuleCount
End Property

Property Let SpecialModuleCount(NewVal As Long)
    pSpecialModuleCount = NewVal
    ReDim Preserve pSpecialModules(pSpecialModuleCount)
End Property

Property Get SpecialModules(Index As Long) As VBSpecialModule
    Set SpecialModules = pSpecialModules(Index)
End Property

Property Set SpecialModules(Index As Long, NewObj As VBSpecialModule)
    Set pSpecialModules(Index) = NewObj
End Property


Property Get StringCount() As Long
    StringCount = pStringCount
End Property

Property Let StringCount(NewValue As Long)
    ReDim Preserve pStrings(NewValue)
    pStringCount = NewValue
End Property

Property Get Strings(Index As Long) As VBString
    Set Strings = pStrings(Index)
End Property

Property Set Strings(Index As Long, NewValue As VBString)
    Set pStrings(Index) = NewValue
End Property


Property Get StringIndex(Str As String) As Long
    
    Dim i As Long
    
    For i = 1 To StringCount
        If Strings(i).Value = Str Then
            StringIndex = i
            Exit Property
        End If
    Next i
    
End Property


Function LoadProjectFile(ProjectFilePath As String)
    
    Dim i As Long
    Dim c As String
    Dim x As Long
    
    Dim ValueName As String
    Dim Value As String
    Dim ModPath As String
    
    Dim NewLineID As Long
    Dim NewLine As String
    Dim NewLine2 As String
    Dim TmpPath As String
    
    Dim NewLineCount As Long
    Dim NewLines() As String
    
    Class_Initialize
    
    With Alias
         Set .ParentProject = Me
         .Scope = VBLocal + VBPrivate + VBPublic
        .AliasType = VBProject
    End With
    
    FilePath = ProjectFilePath
    Code = ReadFile(FilePath)
    
    pLines = Split(Code, vbNewLine)
    
    
    

    'Debug.Print NewLine
    ModuleCount = 0
    ReDim pModules(0)
    
    
    
    For i = 0 To UBound(pLines())
        If InStr(1, Lines(i), "=") <> 0 Then
            ValueName = Left(Lines(i), InStr(1, Lines(i), "=") - 1)
            Value = Right(Lines(i), Len(Lines(i)) - InStr(1, Lines(i), "="))
            Select Case ValueName
                
                Case "Name"
                    Name = FilterString(Value, Chr(34))
                    Alias.Name = Name
                    
                    
                Case "Module", "Form", "Class", "UserControl", "PropertyPage", "Designer"
                    
                    ModuleCount = ModuleCount + 1
                    Set Modules(ModuleCount) = New VBModule
                    
                    Modules(ModuleCount).ID = ModuleCount
                    Set Modules(ModuleCount).ParentProject = Me
                    Modules(ModuleCount).DataType = ValueName
                    If InStr(1, Value, ";") = 0 Then
                        ModPath = Value
                    Else
                        ModPath = Mid(Value, InStr(1, Value, ";") + 2)
                    End If
                    If InStr(1, ModPath, ":\") = 0 Then
                        ModPath = BaseDirPath & ModPath
                    End If
                    Modules(ModuleCount).FilePath = ModPath
                
                Case "ResFile32"
                    
                    pResFilePath = BaseDirPath & GetFileName(FilterString(Value, Chr(34)))
                    ResFileAlias.Name = GetFileName(pResFilePath)
                    
                Case "ExeName32"
                    pExeName = FilterString(Value, Chr(34))
                    
            End Select
        End If
    Next i
    
    For i = 1 To ModuleCount
            
        Set CurrentModule = Modules(i)
        Modules(i).LoadModuleFile Modules(i).FilePath
    
    Next i
    
    'FileCopy TmpPath, "D:\dec.bas"
    If FileExist(TmpPath) Then Kill TmpPath
    
    'EncryptStrings
    
    
End Function


Sub AddModule(ModStr, DataType As String, SrcPath As String)
    
    Dim sName As String
    Dim sPath As String
    
    If InStr(1, ModStr, "; ") <> 0 Then
        sName = Left(ModStr, InStr(1, ModStr, "; ") - 1)
        sPath = BaseDirPath & Right(ModStr, Len(ModStr) - InStr(1, ModStr, "; ") - 1)
    Else
        sName = Left(ModStr, Len(ModStr) - 4)
        sPath = BaseDirPath & ModStr
    End If
    
    If InStr(1, ModStr, Chr(34)) <> 0 Then
        sName = Mid(ModStr, 2, Len(ModStr) - 2)
        sPath = BaseDirPath & sName
    End If
    
    
    ModuleCount = ModuleCount + 1
    
    Set pModules(ModuleCount) = New VBModule
    
    With pModules(UBound(pModules))
        .ID = ModuleCount
        .FilePath = SrcPath
        .DataType = DataType
        .LoadModuleFile sPath
    End With
    
End Sub



Function Copy(DstDir As String)
    Dim i As Long
    Dim NewName As String
    
    Dim SrcPath As String
    Dim DstPath As String
    
    Dim ModuleBin As String
    Dim ProjectBin As String
    
    Cancelled = False
    
    If DirExist(DstDir) Then
        'If MsgBox("Das Zielverzeichnis " & DstDir & " existiert bereits. Soll es überschrieben werden?", vbQuestion + vbYesNo) = vbYes Then
            DeleteDir DstDir
        'Else
        '    Exit Function
        'End If
    End If
    
    If Not BuildPath(DstDir) Then
        MsgBox "Das Zielverzeichnis konnte nicht erstellt werden.", vbCritical
        Exit Function
    End If
    
    ProjectBin = ReadFile(FilePath)
    
    For i = 1 To ModuleCount
        If Cancelled Then Exit Function
        
        SrcPath = ""
        
        ModuleBin = ReadFile(Modules(i).FilePath)
        'ProjectBin = Replace(ProjectBin, Modules(i).Name, Modules(i).FileName)
        
            
        SrcPath = Modules(i).FilePath
        DstPath = FixPath(DstDir) & Modules(i).FileName
        
        FileCopy SrcPath, DstPath
        
        SrcPath = Modules(i).ResFilePath
        DstPath = FixPath(DstDir) & Modules(i).ResFileName
        
        If FileExist(SrcPath) Then
            FileCopy SrcPath, DstPath
        End If
        
        
        'If Modules(i).ResFileName <> "" Then
        '    ModuleBin = Replace(ModuleBin, Modules(i).ResFileName, Modules(i).ResFileName)
        'End If
        
        WriteFile FixPath(DstDir) & Modules(i).FileName, ModuleBin
        
        RaiseEvent ProjectProgress(100 / ModuleCount * i)
    Next i
    
    WriteFile FixPath(DstDir) & FileName, ProjectBin
    
    RaiseEvent ProjectProgress(100)
    
End Function

Function PreProcess()
    
    Dim i As Long
    
    
    AddSpecialModules
    
    
    For i = 1 To ModuleCount
        
        Modules(i).PreProcess
        
    Next i
    
    'pOutDir = FixPath(My.Paths.AppPath)
    
    If My.Config.ExportStrings Or My.Config.EnableCrackCheck Then
        SaveStrings pOutDir
    End If
    
    RaiseEvent ProjectProgress(100)
    
End Function

Function CryptProject(DestinationDir As String)
    
    Dim i As Long
    

    
    For i = 1 To ModuleCount
      
        Set CurrentModule = Modules(i)
        
        
        
        With CurrentModule
            
            
            If My.Config.Obfuscate Then
                .CryptModule
            End If
            
            'Me.Code = ReplaceQuotations(Me.Code, .FileAlias.Name, .FileAlias.NewName)
            'Me.Code = ReplaceQuotations(Me.Code, .Alias.Name, .Alias.NewName)
            
            'Debug.Print Modules(i).Lines(46)
            
            ReplaceInLines pLines, 0, UBound(pLines), .FilePath, .FileAlias.NewName
            
            ReplaceInLines pLines, 0, UBound(pLines), Quote(.FileAlias.Name), Quote(.FileAlias.NewName)
            ReplaceInLines pLines, 0, UBound(pLines), .FileAlias.Name, .FileAlias.NewName
            
            ReplaceInLines pLines, 0, UBound(pLines), Quote(.Alias.Name), Quote(.Alias.NewName)
            ReplaceInLines pLines, 0, UBound(pLines), .Alias.Name, .Alias.NewName
            
        
        End With
        
        If Cancelled Then Exit Function
        
    Next i
    
    If ResFilePath <> "" Then
        ReplaceInLines pLines, 0, UBound(pLines), Quote(ResFileAlias.Name), Quote(ResFileAlias.NewName)
        ReplaceInLines pLines, 0, UBound(pLines), ResFileAlias.Name, ResFileAlias.NewName
    End If
    
    RaiseEvent ProjectProgress(100)
    
    
End Function

Function Compile()
    
    Dim Reg As MyRegistryClass
  
    Dim VBEXEPath As String
    Dim cmd As String
    
    
    Set Reg = New MyRegistryClass
    With Reg
        .ClassKey = HKEY_CLASSES_ROOT
        .SectionKey = "Applications\vb6.exe\shell\Make\command"
        If CStr(.Value) = "" Then
            MsgBox "Visual Basic 6 seems to be not installed.", vbCritical
            Exit Function
        End If
        VBEXEPath = Left(.Value, InStr(1, .Value, """%1""") - 2)
    End With
    Set Reg = Nothing
    
    cmd = VBEXEPath & " " & GetShortName(FilePath) & " /make /outdir " & GetShortName(pOutDir)
    
    Run cmd, True 'compile EXE
    
    Debug.Print cmd
    
    If My.Config.EnableCrackCheck Then
      
      CryptStrRes
        
    End If
    
End Function

Function CryptStrRes()

    Dim ExePath As String
    Dim StrRes As String
    Dim CRC As String
    Dim EncStrFilePath As String
        
    ExePath = OutDir & ExeName
    EncStrFilePath = OutDir & ExeName & ".res"
        
    CRC = CStr(GetCRC32FromFile(ExePath))
    
    StrRes = ReadFile(StringResFilePath)
    StrRes = RC4(StrRes, CRC)
    
    WriteFile EncStrFilePath, StrRes

End Function

Private Function AddSpecialModule(Name As String, FilePath As String, DataType As String, Optional ResFilePath As String = "")
    
    Dim SpecialModule As VBSpecialModule
    
    SpecialModuleCount = SpecialModuleCount + 1
    
    Set SpecialModules(SpecialModuleCount) = New VBSpecialModule
    
    Set SpecialModule = Me.SpecialModules(SpecialModuleCount)
    
    
    With SpecialModule
        
        .ID = SpecialModuleCount
        .Name = Name
        .FilePath = FilePath
        .ResFilePath = ResFilePath
        .Code = ReadFile(FilePath)
        .DataType = DataType
        
        
        Select Case .DataType
            
            Case "Module"
                
                .ProjectLine = "Module=" & .Name & "; " & GetFileName(.FilePath)
                
            Case "Form"
                
                .ProjectLine = "Form=" & GetFileName(.FilePath)
            
            Case Else
            
                Debug.Assert False
            
        End Select
        
        
    End With
    
End Function


Function AddSpecialModules()
    
    Dim i As Long
    
    Dim ValueName As String
    Dim Value As String
    
    Dim NewLineID As Long
    
    'Dim NewLineCount As Long
    'Dim NewLines() As String
    
    If My.Config.Translate Then
        
'        NewLineCount = NewLineCount + 1
'        'SpecialModuleID(TranslationModule) = NewLineCount
'        ReDim Preserve NewLines(1 To NewLineCount)
'        NewLines(NewLineCount) = "Module=VBTranslationModule; " & "VBTranslationModule.bas"
'
'        NewLineCount = NewLineCount + 1
'        'SpecialModuleID(TranslationForm) = NewLineCount
'        ReDim Preserve NewLines(1 To NewLineCount)
'        NewLines(NewLineCount) = "Form=" & "VBTranslationForm.frm"
    
        AddSpecialModule "VBTranslationModule", FixPath(App.Path) & "VBTranslationModule.bas", "Module"
        AddSpecialModule "VBTranslationForm", FixPath(App.Path) & "VBTranslationForm.frm", "Form", FixPath(App.Path) & "VBTranslationForm.frx"
        
    End If
    
    
    For i = 0 To UBound(pLines())
    
        If InStr(1, Lines(i), "=") <> 0 Then
        
            ValueName = Left(Lines(i), InStr(1, Lines(i), "=") - 1)
            Value = Right(Lines(i), Len(Lines(i)) - InStr(1, Lines(i), "="))
            
            Select Case ValueName
                
                Case "Reference", "Object"
                    
                    'AddLib Lines(i)
                    
                Case "Module", "Form", "Class", "UserControl", "PropertyPage", "Designer"
                    
                    NewLineID = i
                    Exit For
                
                Case Else
                    
            End Select
        End If
        
    Next i
    
    If NewLineID <> 0 Then
    
        For i = 1 To SpecialModuleCount
        
            AddLine NewLineID + (i - 1), Me.SpecialModules(i).ProjectLine
            
        Next i
        
    End If
    
    
End Function

Function Rebuild() As String
    
    Rebuild = Join(pLines, vbNewLine)

End Function

Function GlobalReplace(OldWord As String, NewWord As String, Optional ReplaceInProjectFile As Boolean = False)
    
    Dim i As Long
    
    For i = 1 To ModuleCount
        With Modules(i)
            
            '.Code = ReplaceWords(.Code, OldWord, NewWord)
            
            Modules(i).ReplaceInMod OldWord, NewWord
            
            'Modules(i).ReplaceInSubs OldWord, NewWord
        
        End With
    Next i
    
    If ReplaceInProjectFile Then
        ReplaceInLines pLines, 0, UBound(pLines), OldWord, NewWord
    End If
    'Code = ReplaceAlias(Code, Me.Alias)
    
End Function


Private Sub CurrentModule_FileProgress(Percent As Double)
    RaiseEvent FileProgress(Percent)
    RaiseEvent ProjectProgress((100 / ModuleCount) * (CurrentModule.ID - 1 + ((1 / 100) * Percent)))
End Sub


'Function EncryptStrings()
'
'    Dim i As Long
'    Dim x As Long
'
'    Dim Value As String
'    Dim NewValue As String
'
'    For i = 1 To ModuleCount
'
'        For x = 1 To Modules(i).StringCount
'
'            If InStrArray(Modules(i).Strings(x), pStrings) = -1 Then
'
'                StringCount = StringCount + 1
'                Strings(StringCount) = Modules(i).Strings(x)
'
'                Debug.Print Strings(StringCount)
'
'            End If
'
'        Next x
'
'    Next i
'
'End Function



Function CreateStringDecModule() As String
    
'    Dim SrcPath As String
'    Dim DstPath As String
'    Dim FileData As String
'
'    SrcPath = FixPath(App.Path) & "VBStringDecryptionModule.bas"
'    DstPath = TempFilePath(".bas")
'
'    FileData = ReadFile(SrcPath)
'
'    FileData = Replace(FileData, StringEncryption.DefaultBase64Key, StringEncryption.Base64Key)
'
'    WriteFile DstPath, FileData
'
'    CreateStringDecModule = GetLongFilename(DstPath)
    
    
End Function


Function AddLine(LineID As Long, LineStr As String)
    
    Dim i As Long
    Dim c As Long
    
    ReDim Preserve pLines(UBound(pLines) + 1)
    
    c = 0
    Do
    
        i = UBound(pLines) - c
        
        pLines(i) = pLines(i - 1)
        
        c = c + 1
    
    Loop While c <= (UBound(pLines) - LineID)
    
    pLines(LineID) = LineStr
    
    'Debug.Print JoinLines(pLines, 0, UBound(pLines))
    
End Function


Sub SaveStrings(DestinationDir As String)
    
    Dim i As Long
    Dim v(1 To 4) As String
    Dim Bin As String
    Dim StrFilePath As String
    
    Dim MaxLineCount As Long
    
    For i = 1 To ModuleCount
        If Modules(i).LineCount > MaxLineCount Then
            MaxLineCount = Modules(i).LineCount
        End If
    Next i
    
    For i = 1 To StringCount
        
        'v(1) = NullNumber(Strings(i).ParentModule.ID, Strings(i).ParentModule.ParentProject.ModuleCount)
        'v(2) = NullNumber(Strings(i).LineID, MaxLineCount)
        'v(3) = NullNumber(Strings(i).Index, 10)
        'v(4) = "'" & Strings(i).Value & "'"
        
        'Bin = Bin & NullNumber(Strings(i).GlobalID, StringCount) & "=" & Join(v(), "|")
        
        Bin = Bin & Strings(i).Value
        
        If i <> StringCount Then
            Bin = Bin & vbNewLine
        End If
        
    Next i
    
    StrFilePath = FixPath(DestinationDir) & Me.Name & ".str"
    
    If FileExist(StrFilePath) Then
        DeleteFile (StrFilePath)
    End If
    WriteFile StrFilePath, Bin
    
End Sub

Private Function NullNumber(Value As Long, Max As Long) As String
    Dim n As Long
    n = Len(CStr(Max))
    NullNumber = Format(Value, String(n, "0"))
End Function


'Property Get SpecialModuleID(SpecialModule As VBSpecialModuleTypes) As Long
'    SpecialModuleID = pSpecialModuleID(SpecialModule)
'End Property
'
'Property Get SpecialModuleCount() As Long
'    SpecialModuleCount = UBound(pSpecialModuleID)
'End Property
'
'
'Property Let SpecialModuleID(SpecialModule As VBSpecialModuleTypes, NewValue As Long)
'    pSpecialModuleID(SpecialModule) = NewValue
'End Property

Public Property Get OutDir() As String
    OutDir = pOutDir
End Property

Public Property Let OutDir(NewVal As String)
    pOutDir = NewVal
End Property


Public Property Get ExeName() As String
    ExeName = pExeName
End Property






Private Function CalcCRC32(ByteArray() As Byte) As Long
    
    Dim m_CRC32 As Long
    Dim m_CRC32Asm() As Byte
    Dim m_CRC32Table(0 To 255) As Long
    
    Dim i As Long
    Dim sASM As String
  
    Dim ByteSize As Long
    
    m_CRC32Table(0) = &H0
    m_CRC32Table(1) = &H77073096
    m_CRC32Table(2) = &HEE0E612C
    m_CRC32Table(3) = &H990951BA
    m_CRC32Table(4) = &H76DC419
    m_CRC32Table(5) = &H706AF48F
    m_CRC32Table(6) = &HE963A535
    m_CRC32Table(7) = &H9E6495A3
    m_CRC32Table(8) = &HEDB8832
    m_CRC32Table(9) = &H79DCB8A4
    m_CRC32Table(10) = &HE0D5E91E
    m_CRC32Table(11) = &H97D2D988
    m_CRC32Table(12) = &H9B64C2B
    m_CRC32Table(13) = &H7EB17CBD
    m_CRC32Table(14) = &HE7B82D07
    m_CRC32Table(15) = &H90BF1D91
    m_CRC32Table(16) = &H1DB71064
    m_CRC32Table(17) = &H6AB020F2
    m_CRC32Table(18) = &HF3B97148
    m_CRC32Table(19) = &H84BE41DE
    m_CRC32Table(20) = &H1ADAD47D
    m_CRC32Table(21) = &H6DDDE4EB
    m_CRC32Table(22) = &HF4D4B551
    m_CRC32Table(23) = &H83D385C7
    m_CRC32Table(24) = &H136C9856
    m_CRC32Table(25) = &H646BA8C0
    m_CRC32Table(26) = &HFD62F97A
    m_CRC32Table(27) = &H8A65C9EC
    m_CRC32Table(28) = &H14015C4F
    m_CRC32Table(29) = &H63066CD9
    m_CRC32Table(30) = &HFA0F3D63
    m_CRC32Table(31) = &H8D080DF5
    m_CRC32Table(32) = &H3B6E20C8
    m_CRC32Table(33) = &H4C69105E
    m_CRC32Table(34) = &HD56041E4
    m_CRC32Table(35) = &HA2677172
    m_CRC32Table(36) = &H3C03E4D1
    m_CRC32Table(37) = &H4B04D447
    m_CRC32Table(38) = &HD20D85FD
    m_CRC32Table(39) = &HA50AB56B
    m_CRC32Table(40) = &H35B5A8FA
    m_CRC32Table(41) = &H42B2986C
    m_CRC32Table(42) = &HDBBBC9D6
    m_CRC32Table(43) = &HACBCF940
    m_CRC32Table(44) = &H32D86CE3
    m_CRC32Table(45) = &H45DF5C75
    m_CRC32Table(46) = &HDCD60DCF
    m_CRC32Table(47) = &HABD13D59
    m_CRC32Table(48) = &H26D930AC
    m_CRC32Table(49) = &H51DE003A
    m_CRC32Table(50) = &HC8D75180
    m_CRC32Table(51) = &HBFD06116
    m_CRC32Table(52) = &H21B4F4B5
    m_CRC32Table(53) = &H56B3C423
    m_CRC32Table(54) = &HCFBA9599
    m_CRC32Table(55) = &HB8BDA50F
    m_CRC32Table(56) = &H2802B89E
    m_CRC32Table(57) = &H5F058808
    m_CRC32Table(58) = &HC60CD9B2
    m_CRC32Table(59) = &HB10BE924
    m_CRC32Table(60) = &H2F6F7C87
    m_CRC32Table(61) = &H58684C11
    m_CRC32Table(62) = &HC1611DAB
    m_CRC32Table(63) = &HB6662D3D
    m_CRC32Table(64) = &H76DC4190
    m_CRC32Table(65) = &H1DB7106
    m_CRC32Table(66) = &H98D220BC
    m_CRC32Table(67) = &HEFD5102A
    m_CRC32Table(68) = &H71B18589
    m_CRC32Table(69) = &H6B6B51F
    m_CRC32Table(70) = &H9FBFE4A5
    m_CRC32Table(71) = &HE8B8D433
    m_CRC32Table(72) = &H7807C9A2
    m_CRC32Table(73) = &HF00F934
    m_CRC32Table(74) = &H9609A88E
    m_CRC32Table(75) = &HE10E9818
    m_CRC32Table(76) = &H7F6A0DBB
    m_CRC32Table(77) = &H86D3D2D
    m_CRC32Table(78) = &H91646C97
    m_CRC32Table(79) = &HE6635C01
    m_CRC32Table(80) = &H6B6B51F4
    m_CRC32Table(81) = &H1C6C6162
    m_CRC32Table(82) = &H856530D8
    m_CRC32Table(83) = &HF262004E
    m_CRC32Table(84) = &H6C0695ED
    m_CRC32Table(85) = &H1B01A57B
    m_CRC32Table(86) = &H8208F4C1
    m_CRC32Table(87) = &HF50FC457
    m_CRC32Table(88) = &H65B0D9C6
    m_CRC32Table(89) = &H12B7E950
    m_CRC32Table(90) = &H8BBEB8EA
    m_CRC32Table(91) = &HFCB9887C
    m_CRC32Table(92) = &H62DD1DDF
    m_CRC32Table(93) = &H15DA2D49
    m_CRC32Table(94) = &H8CD37CF3
    m_CRC32Table(95) = &HFBD44C65
    m_CRC32Table(96) = &H4DB26158
    m_CRC32Table(97) = &H3AB551CE
    m_CRC32Table(98) = &HA3BC0074
    m_CRC32Table(99) = &HD4BB30E2
    m_CRC32Table(100) = &H4ADFA541
    m_CRC32Table(101) = &H3DD895D7
    m_CRC32Table(102) = &HA4D1C46D
    m_CRC32Table(103) = &HD3D6F4FB
    m_CRC32Table(104) = &H4369E96A
    m_CRC32Table(105) = &H346ED9FC
    m_CRC32Table(106) = &HAD678846
    m_CRC32Table(107) = &HDA60B8D0
    m_CRC32Table(108) = &H44042D73
    m_CRC32Table(109) = &H33031DE5
    m_CRC32Table(110) = &HAA0A4C5F
    m_CRC32Table(111) = &HDD0D7CC9
    m_CRC32Table(112) = &H5005713C
    m_CRC32Table(113) = &H270241AA
    m_CRC32Table(114) = &HBE0B1010
    m_CRC32Table(115) = &HC90C2086
    m_CRC32Table(116) = &H5768B525
    m_CRC32Table(117) = &H206F85B3
    m_CRC32Table(118) = &HB966D409
    m_CRC32Table(119) = &HCE61E49F
    m_CRC32Table(120) = &H5EDEF90E
    m_CRC32Table(121) = &H29D9C998
    m_CRC32Table(122) = &HB0D09822
    m_CRC32Table(123) = &HC7D7A8B4
    m_CRC32Table(124) = &H59B33D17
    m_CRC32Table(125) = &H2EB40D81
    m_CRC32Table(126) = &HB7BD5C3B
    m_CRC32Table(127) = &HC0BA6CAD
    m_CRC32Table(128) = &HEDB88320
    m_CRC32Table(129) = &H9ABFB3B6
    m_CRC32Table(130) = &H3B6E20C
    m_CRC32Table(131) = &H74B1D29A
    m_CRC32Table(132) = &HEAD54739
    m_CRC32Table(133) = &H9DD277AF
    m_CRC32Table(134) = &H4DB2615
    m_CRC32Table(135) = &H73DC1683
    m_CRC32Table(136) = &HE3630B12
    m_CRC32Table(137) = &H94643B84
    m_CRC32Table(138) = &HD6D6A3E
    m_CRC32Table(139) = &H7A6A5AA8
    m_CRC32Table(140) = &HE40ECF0B
    m_CRC32Table(141) = &H9309FF9D
    m_CRC32Table(142) = &HA00AE27
    m_CRC32Table(143) = &H7D079EB1
    m_CRC32Table(144) = &HF00F9344
    m_CRC32Table(145) = &H8708A3D2
    m_CRC32Table(146) = &H1E01F268
    m_CRC32Table(147) = &H6906C2FE
    m_CRC32Table(148) = &HF762575D
    m_CRC32Table(149) = &H806567CB
    m_CRC32Table(150) = &H196C3671
    m_CRC32Table(151) = &H6E6B06E7
    m_CRC32Table(152) = &HFED41B76
    m_CRC32Table(153) = &H89D32BE0
    m_CRC32Table(154) = &H10DA7A5A
    m_CRC32Table(155) = &H67DD4ACC
    m_CRC32Table(156) = &HF9B9DF6F
    m_CRC32Table(157) = &H8EBEEFF9
    m_CRC32Table(158) = &H17B7BE43
    m_CRC32Table(159) = &H60B08ED5
    m_CRC32Table(160) = &HD6D6A3E8
    m_CRC32Table(161) = &HA1D1937E
    m_CRC32Table(162) = &H38D8C2C4
    m_CRC32Table(163) = &H4FDFF252
    m_CRC32Table(164) = &HD1BB67F1
    m_CRC32Table(165) = &HA6BC5767
    m_CRC32Table(166) = &H3FB506DD
    m_CRC32Table(167) = &H48B2364B
    m_CRC32Table(168) = &HD80D2BDA
    m_CRC32Table(169) = &HAF0A1B4C
    m_CRC32Table(170) = &H36034AF6
    m_CRC32Table(171) = &H41047A60
    m_CRC32Table(172) = &HDF60EFC3
    m_CRC32Table(173) = &HA867DF55
    m_CRC32Table(174) = &H316E8EEF
    m_CRC32Table(175) = &H4669BE79
    m_CRC32Table(176) = &HCB61B38C
    m_CRC32Table(177) = &HBC66831A
    m_CRC32Table(178) = &H256FD2A0
    m_CRC32Table(179) = &H5268E236
    m_CRC32Table(180) = &HCC0C7795
    m_CRC32Table(181) = &HBB0B4703
    m_CRC32Table(182) = &H220216B9
    m_CRC32Table(183) = &H5505262F
    m_CRC32Table(184) = &HC5BA3BBE
    m_CRC32Table(185) = &HB2BD0B28
    m_CRC32Table(186) = &H2BB45A92
    m_CRC32Table(187) = &H5CB36A04
    m_CRC32Table(188) = &HC2D7FFA7
    m_CRC32Table(189) = &HB5D0CF31
    m_CRC32Table(190) = &H2CD99E8B
    m_CRC32Table(191) = &H5BDEAE1D
    m_CRC32Table(192) = &H9B64C2B0
    m_CRC32Table(193) = &HEC63F226
    m_CRC32Table(194) = &H756AA39C
    m_CRC32Table(195) = &H26D930A
    m_CRC32Table(196) = &H9C0906A9
    m_CRC32Table(197) = &HEB0E363F
    m_CRC32Table(198) = &H72076785
    m_CRC32Table(199) = &H5005713
    m_CRC32Table(200) = &H95BF4A82
    m_CRC32Table(201) = &HE2B87A14
    m_CRC32Table(202) = &H7BB12BAE
    m_CRC32Table(203) = &HCB61B38
    m_CRC32Table(204) = &H92D28E9B
    m_CRC32Table(205) = &HE5D5BE0D
    m_CRC32Table(206) = &H7CDCEFB7
    m_CRC32Table(207) = &HBDBDF21
    m_CRC32Table(208) = &H86D3D2D4
    m_CRC32Table(209) = &HF1D4E242
    m_CRC32Table(210) = &H68DDB3F8
    m_CRC32Table(211) = &H1FDA836E
    m_CRC32Table(212) = &H81BE16CD
    m_CRC32Table(213) = &HF6B9265B
    m_CRC32Table(214) = &H6FB077E1
    m_CRC32Table(215) = &H18B74777
    m_CRC32Table(216) = &H88085AE6
    m_CRC32Table(217) = &HFF0F6A70
    m_CRC32Table(218) = &H66063BCA
    m_CRC32Table(219) = &H11010B5C
    m_CRC32Table(220) = &H8F659EFF
    m_CRC32Table(221) = &HF862AE69
    m_CRC32Table(222) = &H616BFFD3
    m_CRC32Table(223) = &H166CCF45
    m_CRC32Table(224) = &HA00AE278
    m_CRC32Table(225) = &HD70DD2EE
    m_CRC32Table(226) = &H4E048354
    m_CRC32Table(227) = &H3903B3C2
    m_CRC32Table(228) = &HA7672661
    m_CRC32Table(229) = &HD06016F7
    m_CRC32Table(230) = &H4969474D
    m_CRC32Table(231) = &H3E6E77DB
    m_CRC32Table(232) = &HAED16A4A
    m_CRC32Table(233) = &HD9D65ADC
    m_CRC32Table(234) = &H40DF0B66
    m_CRC32Table(235) = &H37D83BF0
    m_CRC32Table(236) = &HA9BCAE53
    m_CRC32Table(237) = &HDEBB9EC5
    m_CRC32Table(238) = &H47B2CF7F
    m_CRC32Table(239) = &H30B5FFE9
    m_CRC32Table(240) = &HBDBDF21C
    m_CRC32Table(241) = &HCABAC28A
    m_CRC32Table(242) = &H53B39330
    m_CRC32Table(243) = &H24B4A3A6
    m_CRC32Table(244) = &HBAD03605
    m_CRC32Table(245) = &HCDD70693
    m_CRC32Table(246) = &H54DE5729
    m_CRC32Table(247) = &H23D967BF
    m_CRC32Table(248) = &HB3667A2E
    m_CRC32Table(249) = &HC4614AB8
    m_CRC32Table(250) = &H5D681B02
    m_CRC32Table(251) = &H2A6F2B94
    m_CRC32Table(252) = &HB40BBE37
    m_CRC32Table(253) = &HC30C8EA1
    m_CRC32Table(254) = &H5A05DF1B
    m_CRC32Table(255) = &H2D02EF8D
    
    sASM = "5589E557565053518B45088B008B750C8B7D108B4D1431DB8A1E30C3C1E80833049F464975F28B4D088901595B585E5F89EC5DC21000"
    
    ReDim m_CRC32Asm(0 To Len(sASM) \ 2 - 1)
    
    For i = 1 To Len(sASM) Step 2
        m_CRC32Asm(i \ 2) = Val("&H" & Mid$(sASM, i, 2))
    Next
    
    On Local Error GoTo NoData
    
    ByteSize = UBound(ByteArray) - LBound(ByteArray) + 1
    
    On Local Error GoTo 0
   
    Call CallWindowProc(VarPtr(m_CRC32Asm(0)), VarPtr(m_CRC32), VarPtr(ByteArray(LBound(ByteArray))), VarPtr(m_CRC32Table(0)), ByteSize)
    
NoData:

    CalcCRC32 = (Not m_CRC32)
    
End Function


Private Function GetCRC32FromFile(FilePath As String) As Long

    Dim Filenr As Integer
    Dim ByteArray() As Byte
    
    If (FileLen(FilePath) = 0) Then Exit Function
    
    Filenr = FreeFile
    Open FilePath For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr
    
    GetCRC32FromFile = CalcCRC32(ByteArray)

End Function


Property Get StringResFilePath()
    StringResFilePath = FixPath(OutDir) & Me.Name & ".str"
End Property


Private Function AddLib(Lib As String)
    
    Dim s() As String
    Dim StrLib As String
    Dim StrID As String
    
    Dim TLI As TypeLibInfo
    
    'Debug.Print Lib
    
    If Left(Lib, Len("Reference")) = "Reference" Then
        
        s = Split(Lib, "#")
        StrLib = s(3)
        
        StrID = s(0)
        StrID = Mid(StrID, InStr(1, StrID, "{"))
        
    ElseIf Left(Lib, Len("Object")) = "Object" Then
        
        s = Split(Lib, "; ")
        StrLib = s(1)
        
        s = Split(Lib, "#")
        
        StrID = s(0)
        StrID = Mid(StrID, InStr(1, StrID, "{"))
        
    End If
    
    Libs.Add StrLib
    LibIDs.Add StrID
    
'    If FileExist(StrLib) Then
'        fObjects.AddFromFile StrLib
'    Else
'        fObjects.AddFromID StrID
'    End If
    
    
End Function


Public Function GetModuleIndex(ModuleName As String) As Long
    
    
    Dim i As Long
    
    For i = 1 To Me.ModuleCount
        
        If Modules(i).Name = ModuleName Then
            
            GetModuleIndex = i
            Exit Function
        
        End If
        
    Next i
    
End Function

Public Function GetStringIndex(StringValue As String) As Long
    
    Dim i As Long
    
    For i = 1 To StringCount
        
        If Strings(i).Value = StringValue Then
            
            GetStringIndex = i
            Exit Function
            
        End If
        
    Next i
    
End Function


Function CreatTranslationINI(FilePath As String)
    
    Dim i As Long
    
    Dim FileData As String
    
    For i = 1 To StringCount
        
        If Strings(i).Translation <> "" Then
            
            SetIniValue FileData, "Translation", Base64Encode(Strings(i).Value), Base64Encode(Strings(i).Translation)
            
        End If
        
    Next i
    
    If FileExist(FilePath) Then
        DeleteFile FilePath
    End If
    
    WriteFile FilePath, FileData
    
End Function
