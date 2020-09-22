VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Paths As MyPaths
Public Config As MyConfig
Public KeyWords As MyKeyWords


Private pAliasCount     As Long
Private pAlias()        As VBAlias

Private WithEvents pProject As VBProject
Attribute pProject.VB_VarHelpID = -1

Private Cancelled As Boolean

Public Event FileProgress(Progress As Double)
Public Event ProjectProgress(Progress As Double)

Private Sub Class_Initialize()
    
    Set Paths = New MyPaths
    Set Config = New MyConfig
    Set KeyWords = New MyKeyWords
    
    ReDim pAlias(0)
    
    Set pProject = New VBProject
    
    'VBProtectedPropertiesString = fOptions.txtProperties  'ReadFile(FixPath(App.Path) & "VBProtectedProperties.txt")
    'VBProtectedPropertiesString = Replace(VBProtectedPropertiesString, vbNewLine, "")
    'VBProtectedProperties() = Split(VBProtectedPropertiesString, " ")
    
End Sub


Private Sub Class_Terminate()
    
    Set Config = Nothing
    
End Sub


Sub Start()

    Dim x As Long
    Dim a As Long
    Dim z As Long
    Dim s As String
    Dim ProjectPath As String
    Dim OutDir As String
    
    
    Config.LoadConfig
    
    
    
    If Command <> "" Then
        
        x = InStr(1, Command, "/Project=", vbTextCompare)
          
        If x <> 0 Then
    
            a = InStr(x + 1, Command, Chr(34))
            z = InStr(a + 1, Command, Chr(34))
            
            ProjectPath = Mid(Command, (a + 1), ((z - 1) - a))
            
            
            x = InStr(1, Command, "/OutDir=", vbTextCompare)
            
            If x <> 0 Then
                
                a = InStr(x + 1, Command, Chr(34))
                z = InStr(a + 1, Command, Chr(34))
            
                OutDir = Mid(Command, (a + 1), ((z - 1) - a))
                
            End If
            
            If InStr(1, Command, "/q", vbTextCompare) = 0 Then
                fProgress.Show
            End If
            
            My.Project.LoadProjectFile ProjectPath
            My.SaveProject OutDir
            
        End If
        
        End
        
    Else
        
        'Load fOptions
        'Load fOpen
        'Load fProgress
        
        fMain.Show
        
    End If
    
End Sub

Property Get Project() As VBProject
    Set Project = pProject
End Property

Property Get AliasCount() As Long
    AliasCount = pAliasCount
End Property

Property Let AliasCount(NewValue As Long)
    ReDim Preserve pAlias(NewValue)
    pAliasCount = NewValue
End Property

Property Get Alias(Index As Long) As VBAlias
    Set Alias = pAlias(Index)
End Property

Property Set Alias(Index As Long, NewObject As VBAlias)
    Set pAlias(Index) = NewObject
End Property


Private Property Get AliasExists(Name As String, Optional AliasType As VBAliasTypes, Optional Scope As VBScopes, Optional ModuleID As Long, Optional SubID As Long) As Long

    Dim i As Long
    Dim b1 As Boolean
    Dim b2 As Boolean
    Dim b3 As Boolean
    Dim b4 As Boolean
    Dim b5 As Boolean
    
    For i = 1 To AliasCount
        
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

Private Property Get NewAliasExists(Name As String, Optional AliasType As VBAliasTypes, Optional Scope As VBScopes, Optional ModuleID As Long, Optional SubID As Long) As Long
    
    Dim r As Long
    
    
    r = AliasExists(Name, AliasType, Scope, ModuleID, SubID)
    If r <> 0 Then
        If pAlias(r).NewName <> "" Then
            NewAliasExists = r
        End If
    End If
    
End Property

Private Property Get AliasFromName(Name As String, Optional AliasType As VBAliasTypes, Optional Scope As VBScopes, Optional ModuleID As Long, Optional SubID As Long) As VBAlias
    Dim r As Long
    r = AliasExists(Name, AliasType, Scope, ModuleID, SubID)
    If r <> 0 Then
        AliasFromName = pAlias(r)
        Exit Property
    End If
End Property

'Private Function AddAlias(Name As String, AliasType As VBAliasTypes, Scope As VBScopes, Optional ModuleID As Long, Optional SubID As Long, Optional VariableID As Long)
'
'    Dim r As Long
'
'    r = AliasExists(Name, AliasType, Scope, ModuleID, SubID)
'    If r <> 0 Then
'        Exit Function
'    End If
'
'    AliasCount = AliasCount + 1
'    With pAlias(AliasCount)
'        .Name = Name
'        .AliasType = AliasType
'        .Scope = Scope
'        .ModuleID = ModuleID
'        .SubID = SubID
'        .VariableID = VariableID
'   End With
'
'End Function


Private Function AddAlias(NewAlias As VBAlias)
    
    Dim r As Long
    Static c As Long
    
    With NewAlias
        
        If .Name = "" Then
            Debug.Assert False
        End If
        
'        If .Name = "VBAlias" Then
'            Debug.Assert False
'        End If
        
        'If InText(1, VBKeyWordsString, .Name) <> 0 Then
        '    .Protected = True
        'End If
        
        
        
        'r = AliasExists(.Name, .AliasType, .Scope, .ModuleID, .SubID)
        'If r <> 0 Then
        '    Debug.Assert False
        '    Exit Function
        'End If
        
    End With
    
    AliasCount = AliasCount + 1
    Set pAlias(AliasCount) = NewAlias
    'pAlias(AliasCount) = AliasCount
    
End Function



Private Property Get AliasUsed(Name As String) As Boolean

    Dim i As Long
    
'    If InStr(1, VBKeyWordsString, Name, vbTextCompare) <> 0 Then
'        AliasUsed = True
'        Exit Property
'    End If

    For i = 1 To pAliasCount
        If (pAlias(i).Name = Name) Or (pAlias(i).NewName = Name) Then
            AliasUsed = True
            Exit Property
        End If
    Next i

End Property



Private Function CreateNewAlias(AliasObj As VBAlias) As String
    
    Const NewLen As Integer = 8
    
    Dim SisterSubs() As VBSub
    Dim s As String
    Dim OldName As String
    Dim NewName As String
    Dim RndName As String
    Dim c As Long
    Dim r As Long
    Dim i As Long
    
    
    If Not Config.Obfuscate Then
        CreateNewAlias = AliasObj.Name
        Exit Function
    End If
    
    
    With AliasObj
        
        If Me.KeyWords.Exists(.Name) Then
            
            CreateNewAlias = .Name
            'Debug.Assert .Name <> "AliasType"
            Exit Function
            
        End If
        
        If .Protected Then
            CreateNewAlias = AliasObj.Name
            Exit Function
        End If
        
        If CBool((.AliasType And VBParameter)) Or CBool((.AliasType And VBSub)) Then
            If .SubID <> 0 Then
                If InStr(1, Project.Modules(.ModuleID).Subs(.SubID).SubType, "Property") <> 0 Then
                
                    'Debug.Print .Name
                
                    c = GetSisterSubs(.ParentSub, SisterSubs())
                
                    For i = 1 To c
                        r = NewAliasExists(.Name, .AliasType, .Scope, .ModuleID, SisterSubs(i).Alias.SubID)
                        If r <> 0 Then
                            'Debug.Assert False
                            CreateNewAlias = pAlias(r).NewName
                            Exit Function
                        End If
                    Next i
                
                End If
            End If
        End If
    
        If (.AliasType And VBSub) Then
            
            If LCase(.Name) = "main" Then
                CreateNewAlias = .Name
                Exit Function
            End If
            
            If .ParentSub.IsEventHandler Then
                
                CreateNewAlias = CreateNewEventHandlerAlias(AliasObj)
                Exit Function
                
            End If
            
        End If
    
        'r = NewAliasExists(.Name, .AliasType, .Scope, .ModuleID, .SubID)
        r = NewAliasExists(.Name)
        If r <> 0 Then
            CreateNewAlias = pAlias(r).NewName
            Exit Function
        End If
    
        
        RndName = RandomAlias(NewLen)
        NewName = RndName
        Do While AliasUsed(NewName)
            c = c + 1
            If c <= 3 Then
                NewName = RandomAlias(NewLen)
            Else
                NewName = RndName & CStr((c - 2))
            End If
        Loop
        
        NewName = "a" & .Name & "Z"
        
        If (.AliasType And vbFile) Then
            NewName = NewName & Right(.Name, 4)
        End If
    
    End With
    
    CreateNewAlias = NewName
    
    

        
    
End Function

Function CreateNewEventHandlerAlias(AliasObj As VBAlias) As String
    
    Dim NewAlias As String
    Dim OldEventSrc As String
    Dim OldEventName As String
    Dim OldEventSrcType As String
    Dim EventSrcMod As VBModule
    Dim NewEventSrc As String
    Dim NewEventName As String
    Dim NewSubName As String
    
    Dim r As Long
    Dim i As Long
    
    With AliasObj
    
        Debug.Assert (.Name <> "ProgressBar_MouseDown")
                    
        OldEventSrc = Left(.Name, InStr(1, .Name, "_") - 1)
        OldEventName = Mid(.Name, InStrRev(.Name, "_") + 1)
        
        r = NewAliasExists(OldEventSrc, VBControl Or VBVariable, 0, .ModuleID)
                    
        If r <> 0 Then
                        
            'r = newaliasexists(.ParentModule.Subs(.SubID)
                    
            NewEventSrc = pAlias(r).NewName
            
            If (pAlias(r).VariableID <> 0) Then
                OldEventSrcType = pAlias(r).ParentVariable.DataType
            ElseIf (pAlias(r).ControlID <> 0) Then
                OldEventSrcType = pAlias(r).ParentControl.DataType
                If InStrRev(OldEventSrcType, ".") <> 0 Then
                    OldEventSrcType = Mid(OldEventSrcType, InStrRev(OldEventSrcType, ".") + 1)
                End If
            End If
            
            If OldEventSrcType <> "" Then
            
                r = NewAliasExists(OldEventSrcType, VBModule, 0)
                If r <> 0 Then
                    Set EventSrcMod = pAlias(r).ParentModule
                    For i = 1 To EventSrcMod.EventCount
                        
                        If EventSrcMod.Events(i).Name = OldEventName Then
                            
                            NewEventName = EventSrcMod.Events(i).Alias.NewName
                            Exit For
                        End If
                        
                    Next i
                End If
                
            End If
            
            
            
            If NewEventName <> "" Then
            
                NewSubName = NewEventSrc & "_" & NewEventName
            
            Else
            
            
                NewSubName = NewEventSrc & "_" & Mid(.Name, InStrRev(.Name, "_") + 1)
            
            End If
            
        Else
        
            NewSubName = .Name
            
        End If
    
    End With
    
    CreateNewEventHandlerAlias = NewSubName
    
End Function


Function AddAliases()
    
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim z As Long
    
    'Add Project (File-)Names
    AddAlias Project.Alias
    AddAlias Project.FileAlias
    
    If Project.ResFileName <> "" Then
        AddAlias Project.ResFileAlias
    End If
    
    'Add Module (File-)Names
    For i = 1 To Project.ModuleCount
        With Project.Modules(i)
            'Debug.Print .DataType, .ID, .Name
            AddAlias .Alias
            AddAlias .FileAlias
            If .ResFileName <> "" Then
                AddAlias .ResFileAlias
                'Debug.Print .DataType, .ID, .ResFileAlias.Name
            End If
        End With
    Next i
    
    'Add Module.Controls
    For x = 1 To Project.ModuleCount
        For y = 1 To Project.Modules(x).ControlCount
            With Project.Modules(x).Controls(y)
                'Debug.Print .DataType, .ID, .Name
                AddAlias .Alias
            End With
        Next y
    Next x
    
    'Add Module.Declarations
    For x = 1 To Project.ModuleCount
        For y = 1 To Project.Modules(x).DeclarationCount
            With Project.Modules(x).Declarations(y)
                'Debug.Print .DataType, .ID, .Name
                AddAlias .Alias
                For z = 1 To .ParameterCount
                    AddAlias .Parameters(z).Alias
                Next z
            End With
        Next y
    Next x
    
    'Add Module.Events
    For x = 1 To Project.ModuleCount
        For y = 1 To Project.Modules(x).EventCount
            With Project.Modules(x).Events(y)
                'Debug.Print .DataType, .ID, .Name
                AddAlias .Alias
                For z = 1 To .ParameterCount
                    AddAlias .Parameters(z).Alias
                Next z
            End With
        Next y
    Next x
    
    'Add Module.Types
    For x = 1 To Project.ModuleCount
        For y = 1 To Project.Modules(x).TypeCount
            With Project.Modules(x).Types(y)
                'Debug.Print .DataType, .ID, .Name
                AddAlias .Alias
                For z = 1 To .MemberCount
                    AddAlias .Members(z).Alias
                Next z
            End With
        Next y
    Next x
    
    'Add Module.Enums
    For x = 1 To Project.ModuleCount
        For y = 1 To Project.Modules(x).EnumCount
            With Project.Modules(x).Enums(y)
                'Debug.Print .DataType, .ID, .Name
                AddAlias .Alias
                For z = 1 To .ValueCount
                    AddAlias .Values(z).Alias
                Next z
            End With
        Next y
    Next x
    
    'Add Module.Variables
    For x = 1 To Project.ModuleCount
        For y = 1 To Project.Modules(x).VariableCount
            With Project.Modules(x).Variables(y)
                'Debug.Print .DataType, .ID, .Name
                AddAlias .Alias
            End With
        Next y
    Next x
    
    'Add Module.Subs
    For x = 1 To Project.ModuleCount
        For y = 1 To Project.Modules(x).SubCount
            With Project.Modules(x).Subs(y)
                'Debug.Print .DataType, .ID, .Name, .IsEventHandler
                AddAlias .Alias
            End With
        Next y
    Next x
    
    'Add Module.Subs.Parameters
    For x = 1 To Project.ModuleCount
        For y = 1 To Project.Modules(x).SubCount
            For z = 1 To Project.Modules(x).Subs(y).ParameterCount
                With Project.Modules(x).Subs(y).Parameters(z)
                    'Debug.Print .DataType, .ID, .Name
                    AddAlias .Alias
                End With
            Next z
        Next y
    Next x
    
    'Add Module.Subs.Variables
     For x = 1 To Project.ModuleCount
        For y = 1 To Project.Modules(x).SubCount
            For z = 1 To Project.Modules(x).Subs(y).VariableCount
                With Project.Modules(x).Subs(y).Variables(z)
                    'Debug.Print .DataType, .ID, .Name
                    AddAlias .Alias
                End With
            Next z
        Next y
    Next x

End Function

Function SaveProject(DstDir As String)
    
    Dim i As Long
    
    Dim x As Long
    Dim y As Long
    Dim z As Long
    
    Dim NewName As String
    Dim NewFileName As String
    Dim NewFilePath As String
    
    
    LoadKeyWords
    
    Cancelled = False
    
    
    AddAliases
    
    
    AssignNewNames VBControl, VBLocal + VBPrivate + VBPublic
    
    For i = 1 To AliasCount
        With pAlias(i)
           
           If .NewName = "" Then
           
                .NewName = CreateNewAlias(pAlias(i))
                
            End If
            
        End With
    Next i
    
    
    
    If DirExist(DstDir) Then
    '    If MsgBox("The destination folder " & DstDir & " does already exist. Overwrite?", vbQuestion + vbYesNo) = vbYes Then
            If Not DeleteDir(DstDir) Then
                MsgBox "Could not delete dir: " & DstDir, vbCritical
                Exit Function
            End If
            
    '    Else
    '        Exit Function
    '    End If
    End If
    
    If Not BuildPath(DstDir) Then
        MsgBox "Could not create destination folder.", vbCritical
        Exit Function
    End If
    
    Project.OutDir = DstDir
    
    Project.PreProcess
    
'    If Config.Obfuscate Then
'        Project.CryptProject (DstDir)
'    End If
    
    
    
    For i = 1 To Project.ModuleCount
        
        NewFileName = Project.Modules(i).Name & Right(Project.Modules(i).FileAlias.NewName, 4)
        
        Debug.Print NewFileName
        
        NewFilePath = FixPath(DstDir) & NewFileName
        
        WriteFile NewFilePath, Project.Modules(i).FileData
        
        If FileExist(Project.Modules(i).ResFilePath) Then
        
            NewFileName = Project.Modules(i).ResFileAlias.NewName
            NewFilePath = DstDir & NewFileName
            
            FileCopy Project.Modules(i).ResFilePath, NewFilePath
            
        End If
        
        RaiseEvent ProjectProgress(100 / Project.ModuleCount * i)
        
        DoEvents
        
    Next i
    
    For i = 1 To Project.SpecialModuleCount
        
        NewFilePath = DstDir & GetFileName(Project.SpecialModules(i).FilePath)
        WriteFile NewFilePath, Project.SpecialModules(i).Code
        
        If Project.SpecialModules(i).ResFilePath <> "" Then
            
            NewFilePath = DstDir & GetFileName(Project.SpecialModules(i).ResFilePath)
            FileCopy Project.SpecialModules(i).ResFilePath, NewFilePath
            
        End If
        
    Next i
    
    
    NewFilePath = DstDir & Project.FileAlias.NewName
    WriteFile NewFilePath, Project.Rebuild
    
    
    If Project.ResFilePath <> "" Then
        FileCopy Project.ResFilePath, DstDir & Project.ResFileAlias.NewName
    End If
    
    
    Project.CreatTranslationINI DstDir & "Default.lng"
    
End Function

Sub CancelEncryption()
    Cancelled = True
End Sub


   

Private Function GetSisterSubs(SubObj As VBSub, SisterSubs() As VBSub) As Long
    
    Dim i As Long
    Dim c As Long
    
    ReDim SisterSubs(0)
    
    For i = 1 To SubObj.ParentModule.SubCount
        If SubObj.ParentModule.Subs(i).Name = SubObj.Name Then
            c = c + 1
            ReDim Preserve SisterSubs(c)
            Set SisterSubs(c) = SubObj.ParentModule.Subs(i)
        End If
    Next i
    
    GetSisterSubs = c
    
End Function


Private Sub LoadKeyWords()
    
    'VBKeyWordsString = fOptions.txtKeywords
    With Me.KeyWords
    
        .AddList fOptions.txtKeywords.Text
        .AddList fOptions.txtProperties.Text
        .AddList fOptions.txtMiscKeywords.Text
        
        '.AddList fObjects.GetWordList
        
        'Debug.Print .Exists("Version")
        
        'Clipboard.SetText .GetList
        
    End With
    
End Sub


Function AssignNewNames(AliasType As VBAliasTypes, Scopes As VBScopes)
    
    Dim i As Long
    Dim n As Long
    
    For i = 1 To AliasCount
        With pAlias(i)
            If CBool((.AliasType And AliasType)) Then
                If CBool((.Scope And Scopes)) Then
                    .NewName = CreateNewAlias(pAlias(i))
                End If
            End If
        End With
    Next i

End Function

Public Property Get Path() As String
    
    If InIDE Then
        Path = FixPath(App.Path) & "Bin\"
    Else
        Path = FixPath(App.Path)
    End If
    
End Property
