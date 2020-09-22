VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBCodeCryptClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPath As String
Private pName As String
Private pDir As String

Private pAliasCount As Long
Private pAlias() As VBAlias

Private pModuleCount As Long
Private pModules() As VBModule

Private pAttachedFileCount As Long
Private pAttachedFile() As String

Public Event FileProgress(Progress As Double)
Public Event ProjectProgress(Progress As Double)

Public ProjectString As String

Private pCancel As Boolean

Public Property Get Path() As String
    Path = pPath
End Property

Public Property Get RootPath() As String
    RootPath = ExtractDirName(pPath)
End Property

Public Property Let Path(NewValue As String)
    pPath = NewValue
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(NewValue As String)
    pName = NewValue
End Property

Property Get FileName() As String
    FileName = ExtractFileName(Path)
End Property

Property Get FileTitle() As String
    FileTitle = GetFileTitle(Path)
End Property

Private Property Get ModuleCount() As Long
    ModuleCount = pModuleCount
End Property

Private Property Get Modules(Index As Long) As VBModule
    Modules = pModules(Index)
End Property

Property Get ModuleTypeCount(TypeName As String) As Long
    Dim i As Long
    For i = 1 To ModuleCount
        If Modules(i).TypeName = TypeName Then
            ModuleTypeCount = ModuleTypeCount + 1
        End If
    Next i
End Property




Private Property Get AliasCount() As Long
    AliasCount = pAliasCount
End Property

Private Property Let AliasCount(NewValue As Long)
    ReDim Preserve pAlias(NewValue)
    pAliasCount = NewValue
End Property

Private Property Get AliasExists(Name As String, Optional AliasType As VBAliasTypes, Optional Scope As VBAliasScopes, Optional ModuleID As Long, Optional SubID As Long) As Long

    Dim i As Long
    Dim b1 As Boolean
    Dim b2 As Boolean
    Dim b3 As Boolean
    Dim b4 As Boolean
    Dim b5 As Boolean
    
    For i = 1 To pAliasCount
        
        b1 = (pAlias(i).Name = Name)
        If b1 = False Then GoTo Continue
        
        If AliasType = 0 Then
            b2 = True
        Else
            b2 = (pAlias(i).AliasType And AliasType)
        End If
        If b2 = False Then GoTo Continue
        
        If Scope = 0 Then
            b3 = True
        Else
            b3 = (pAlias(i).Scope And Scope)
        End If
        If b3 = False Then GoTo Continue
        
        If ModuleID = 0 Then
            b4 = True
        Else
            b4 = pAlias(i).ModuleID = ModuleID
        End If
        If b4 = False Then GoTo Continue
        
        If SubID = 0 Then
            b5 = True
        Else
            b5 = pAlias(i).SubID = SubID
        End If
        If b5 = False Then GoTo Continue
        
        If (b1 And b2 And b3 And b4 And b5) Then
            AliasExists = i
            Exit Property
        End If
        
Continue:
    Next i

End Property

Private Property Get Alias(Name As String, Optional AliasType As VBAliasTypes, Optional Scope As VBAliasScopes, Optional ModuleID As Long, Optional SubID As Long) As VBAlias
    Dim r As Long
    r = AliasExists(Name, AliasType, Scope, ModuleID, SubID)
    If r <> 0 Then
        Alias = pAlias(r)
        Exit Property
    End If
End Property

Private Function AddAlias(Name As String, AliasType As VBAliasTypes, Scope As VBAliasScopes, Optional ModuleID As Long, Optional SubID As Long, Optional VariableID As Long)
    
    Dim r As Long
    
    r = AliasExists(Name, AliasType, Scope, ModuleID, SubID)
    If r <> 0 Then
        Exit Function
    End If
    
    AliasCount = pAliasCount + 1
    With pAlias(pAliasCount)
        .Name = Name
        .AliasType = AliasType
        .Scope = Scope
        .ModuleID = ModuleID
        .SubID = SubID
        .VariableID = VariableID
    End With
    
End Function

Private Function CreateNewAlias(Name As String, AliasType As VBAliasTypes, Scope As VBAliasScopes, Optional ModuleID As Long, Optional SubID As Long) As String
    
    Const NewLen As Integer = 5
    
    Dim RndName As String
    Dim NewName As String
    Dim c As Long
    Dim x As Long
    Dim r As Long
    
    'r = AliasExists(Name, AliasType, Scope, ModuleID, SubID)
    'If r <> 0 Then
    '    CreateNewAlias = pAlias(r).NewName
    '    Exit Function
    'End If
    
    Randomize Timer
    x = RandomNumber(2, NewLen)
    RndName = RandomName(NewLen - 1)
    RndName = Left(RndName, x - 1) & RandomNumber(0, 9) & Mid(RndName, x)
    NewName = RndName
    Do While AliasExists(NewName, AliasType, Scope, ModuleID, SubID)
        c = c + 1
        NewName = RndName & c
    Loop
    
    If (AliasType And VBFile) Then
        NewName = NewName & Right(Name, 4)
    End If
    
    CreateNewAlias = NewName
    
End Function


Public Function Load(Path As String)
    
    Dim bin As String
    Dim Line() As String
    Dim c As String
    Dim i As Long
    Dim ValueName As String
    Dim Value As String
    
    ReDim pModules(0)
    ReDim pAttachedFile(0)
    ReDim pAlias(0)
    
    pPath = Path
    pDir = ExtractDirName(Path)
    ProjectString = ReadFile(Path)
    Line = Split(ProjectString, vbNewLine)
    ReDim pModules(0)
    For i = 0 To UBound(Line())
        If InStr(1, Line(i), "=") <> 0 Then
            ValueName = Left(Line(i), InStr(1, Line(i), "=") - 1)
            Value = Right(Line(i), Len(Line(i)) - InStr(1, Line(i), "="))
            Select Case ValueName
                Case "Name"
                    pName = FilterString(Value, Chr(34))
                Case "Module", "Form", "Class", "UserControl", "PropertyPage"
                    AddModule GetSrcPath(Value), ValueName
                Case "ResFile32"
                    AddAttachedFile GetSrcPath(Value)
            End Select
        End If
    Next i
    
End Function

Sub AddModule(SrcPath As String, TypeName As String)
    
    Dim FilePath As String

    FilePath = pDir & SrcPath
    
    pModuleCount = pModuleCount + 1
    ReDim Preserve pModules(pModuleCount)
    
    pModules(pModuleCount) = AnalyseModule(ReadFile(FilePath))
    
     With pModules(pModuleCount)
        
        .FilePath = FilePath
        .FileName = ExtractFileName(FilePath)
        .FileExt = GetFileExt(FilePath)
        
        .TypeName = TypeName
        
        .Encrypt = True
        
    End With
    
End Sub


Sub AddAttachedFile(SrcPath As String)
    
    Dim FilePath As String

    FilePath = pPath & SrcPath
    
    pAttachedFileCount = pAttachedFileCount + 1
    ReDim Preserve pAttachedFile(pAttachedFileCount)
    
    pAttachedFile(pAttachedFileCount) = FilePath
    
End Sub


Function Crypt(DstDir As String)
    Dim i As Long
    
    Dim x As Long
    Dim y As Long
    Dim z As Long
    
    Dim NewName As String
    
    Dim Mods() As VBModule
    
    pCancel = False
    
    'If DirExist(DstDir) Then
    '    If MsgBox("The destination folder " & DstDir & " does already exist. Overwrite?", vbQuestion + vbYesNo) = vbYes Then
    '        DeleteDir DstDir
    '    Else
    '        Exit Function
    '    End If
    'End If
    
    If Not BuildPath(DstDir) Then
        MsgBox "Could not create destination folder.", vbCritical
        Exit Function
    End If
    
    ReDim Mods(ModuleCount)
    For i = 1 To ModuleCount
        With Mods(i)
            Mods(i) = AnalyseModule(ReadFile(Modules(i).FilePath))
            .FileName = Modules(i).FileName
            
            AddAlias .Name, VBModule, VBLocal + VBPrivate + VBPublic
            AddAlias .FileName, VBFile, VBLocal + VBPrivate + VBPublic
        End With
    Next i
    
    For x = 1 To ModuleCount
        For y = 1 To Mods(x).SubCount
            For z = 1 To Mods(x).Subs(y).VariableCount
                'Debug.Print Mods(x).Subs(y).Variables(z).Name
                AddAlias Mods(x).Subs(y).Variables(z).Name, Mods(x).Subs(y).Variables(z).AliasType, Mods(x).Subs(y).Variables(z).Scope, x, y, z
            Next z
        Next y
    Next x
    
    For i = 1 To pAliasCount
        
        With pAlias(i)
            If CBool((.AliasType And VBVariable)) And (.Scope = VBLocal) Then
                .NewName = CreateNewAlias(.Name, .AliasType, .Scope, .ModuleID, .SubID)
                Mods(.ModuleID).Subs(.SubID).Variables(.VariableID).NewName = .NewName
                'Debug.Print .Name, .NewName, .SubID
            End If
        End With
    Next i
    
    For i = 1 To ModuleCount
        Mods(i) = CryptModule(Mods(i))
    Next i
    
    Exit Function
    
    For i = 1 To ModuleCount
        If pCancel Then Exit Function
        WriteFile DstDir & Mods(i).FileName, Mods(i).Code
        RaiseEvent ProjectProgress(100 / (ModuleCount + 1) * i)
    Next i
    
    'CryptModule Path, DstDir & ExtractFileName(Path)
    
    'RaiseEvent ProjectProgress(100)
    
End Function

Sub CancelEncryption()
    pCancel = True
End Sub

Function GetSrcPath(src As String) As String
    If InStr(1, src, ";") = 0 Then
        GetSrcPath = src
    Else
        GetSrcPath = Mid(src, InStr(1, src, ";") + 2)
    End If
End Function


Private Function AnalyseModule(Code As String) As VBModule

    Const SubIndicators As String = "Sub|Function|Property"
    
    Dim ThisModule As VBModule
    Dim Encrypt As Boolean
    Dim Lines() As String
    Dim i As Long
    Dim l As Long
    Dim y As Long
    Dim z As Long
    Dim r As Long
    Dim Indicators() As String
    Dim SubStart As Long
    Dim SubEnd As Long
    Dim SubCode As String
    Dim Subs() As VBSub
    
    
    ReDim ThisModule.Subs(0)
    
    r = InText(1, Code, "VB_Name")
    If r <> 0 Then
        ThisModule.Name = GetWord(r + Len("VB_Name"), Code)
    End If
    
    Indicators = Split(SubIndicators, "|")
    Lines = Split(Code, vbNewLine)
    
    l = 0
    Do
        
        i = 0
        Do
        
            r = InText(1, Lines(l), Indicators(i))
            
            If r <> 0 Then
                
                With ThisModule
                
                    SubStart = l
                
                    .SubCount = .SubCount + 1
                    ReDim Preserve .Subs(.SubCount)
                
                    y = l
                    Do
                    
                        r = InText(1, Lines(y), "End " & Indicators(i))
                        If r <> 0 Then
                            SubEnd = y
                            GoTo FoundEnd
                        End If
                    
                        y = y + 1
                    Loop While y <= UBound(Lines)
                
FoundEnd:
                    SubCode = JoinLines(Lines, SubStart, SubEnd)
                    .Subs(.SubCount) = AnalyseSub(SubCode)
                       
                    .Subs(.SubCount).Start = SubStart
                    .Subs(.SubCount).End = SubEnd
                    'Debug.Print .Subs(.SubCount).Start, .Subs(.SubCount).End, .Subs(.SubCount).VariableCount
                
                    l = SubEnd + 1
                    GoTo Continue
                
                End With
                
            End If
            
            i = i + 1
        Loop While i <= UBound(Indicators)
        
Continue:
        
        l = l + 1
    Loop While l <= UBound(Lines)
    
    AnalyseModule = ThisModule
    
End Function

Private Function AnalyseSub(Code As String) As VBSub
    
    Const LocalVarIndicators  As String = "Dim|Static|Const"
    
    Dim ThisSub As VBSub
    
    Dim i As Long
    Dim l As Long
    Dim r As Long
    
    Dim Lines() As String
    Dim Indicators() As String
    
    ReDim Preserve ThisSub.Variables(0)
    ThisSub.Code = Code
    
    Indicators = Split(LocalVarIndicators, "|")
    Lines = Split(Code, vbNewLine)
    
    l = 0
    Do
        
        i = 0
        Do
            
            r = InText(1, Lines(l), Indicators(i))
            If r <> 0 Then
                With ThisSub
                    .VariableCount = .VariableCount + 1
                    ReDim Preserve .Variables(.VariableCount)
                    .Variables(.VariableCount).Name = GetWord(r + Len(Indicators(i)), Lines(l))
                    .Variables(.VariableCount).AliasType = VBVariable
                    .Variables(.VariableCount).Scope = VBLocal
                    'Debug.Print ThisSub.Variables(.VariableCount).Name
                End With
            End If
            
            i = i + 1
        Loop While i <= UBound(Indicators)
    
        l = l + 1
    Loop While l <= UBound(Lines)
    
    AnalyseSub = ThisSub
    
End Function

Private Function CryptModule(OldModule As VBModule) As VBModule
    
    Dim NewModule As VBModule
    Dim i As Long
    
    NewModule = OldModule
    
    NewModule.NewName = CreateNewAlias(OldModule.Name, VBModule, VBLocal + VBPrivate + VBPublic, OldModule.ID)
    
    For i = 1 To OldModule.SubCount
        NewModule.Subs(i) = CryptSub(OldModule.Subs(i))
    Next i
    
    For i = 1 To OldModule.SubCount
        NewModule.Code = Replace(NewModule.Code, OldModule.Subs(i).Code, NewModule.Subs(i).Code)
    Next i
    
    CryptModule = NewModule
    
End Function

   
    
'End Function

Private Function CryptSub(OldSub As VBSub) As VBSub
    
    Dim NewSub As VBSub
    
    Dim i As Long
    Dim l As Long
    Dim r As Long
    Dim s As String
    
    Dim StartPos As Long
    Dim OldName As String
    Dim NewName As String
    
    NewSub = OldSub
    s = OldSub.Code
    
    'Debug.Print s
    
    For i = 1 To OldSub.VariableCount
        
        StartPos = 1
        
        Do
        
            r = InText(StartPos, s, OldSub.Variables(i).Name)
            If r <> 0 Then
                With NewSub
                    StartPos = r
                    
                    'NewName = Alias(OldSub.Variables(i).Name, VBVariable, VBLocal, OldSub.ModuleID, OldSub.ID).NewName
                    s = ReplaceWord(StartPos, s, OldSub.Variables(i).NewName)
                    
                    'Debug.Print s
                    
                    StartPos = r + Len(OldSub.Variables(i).Name)
                    
                End With
            End If
            
        Loop While r <> 0
            
    Next i
    
    'Debug.Print s
    
    NewSub.Code = s
    CryptSub = NewSub
    
End Function


