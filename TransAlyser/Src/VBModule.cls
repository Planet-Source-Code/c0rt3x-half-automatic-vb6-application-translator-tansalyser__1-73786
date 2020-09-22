VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParentProject      As VBProject

Private pFilePath           As String
Private pResFilePath        As String
Private pAlias              As VBAlias
Private pFileAlias          As VBAlias
Private pResFileAlias       As VBAlias

Private pFileData             As String
Private pCode               As String
Private pEncrypt            As Boolean

Private pLineCount          As Long
Private pLines()            As String
Private pLineProtected()    As Boolean
'Private pLineComment()      As String

Private pBlockCount         As Long
Private pBlocks()           As VBCodeBlock

Private pPropertyCount      As Long
Private pProperties()       As VBProperty
    
Private pControlCount       As Long
Private pControls()         As VBControl
    
Private pDeclarationCount   As Long
Private pDeclarations()     As VBDeclaration

Private pEventCount         As Long
Private pEvents()           As VBEvent

Private pEnumCount          As Long
Private pEnums()            As VBEnum

Private pTypeCount          As Long
Private pTypes()            As VBType
    
Private pVariableCount      As Long
Private pVariables()        As VBVariable
    
Private pSubCount           As Long
Private pSubs()             As VBSub

Private pStringCount        As Long
Private pStrings()          As VBString

Private pCaptionCount       As Long
Private pCaptions()         As VBProperty

Private pCaptionsLoaderSubName As String

Private pSubStart           As Long

Public Event FileProgress(Percent As Double)
Public Cancelled As Boolean


Private Sub Class_Initialize()

    Set pAlias = New VBAlias
    Set pFileAlias = New VBAlias
    Set pResFileAlias = New VBAlias
    
    With pAlias
        .AliasType = VBModule
        .Scope = VBLocal + VBPrivate + VBPublic
    End With
    
    With pFileAlias
        .AliasType = vbFile
        .Scope = VBLocal + VBPrivate + VBPublic
    End With
    
    With pResFileAlias
        .AliasType = vbFile
        .Scope = VBPrivate + VBLocal
    End With
    
    ReDim pBlocks(0)
    ReDim pControls(0)
    ReDim pDeclarations(0)
    ReDim pEvents(0)
    ReDim pVariables(0)
    ReDim pSubs(0)
    ReDim pStrings(0)
    ReDim pCaptions(0)
    
End Sub


Property Get ParentProject() As VBProject
    Set ParentProject = pParentProject
End Property

Property Set ParentProject(NewObject As VBProject)
    Set pParentProject = NewObject
    
    Set pAlias.ParentProject = NewObject
    Set pFileAlias.ParentProject = NewObject
    Set pResFileAlias.ParentProject = NewObject
End Property


Property Get ID() As Long
    ID = pAlias.ModuleID
End Property

Property Let ID(NewValue As Long)
    pAlias.ModuleID = NewValue
    pFileAlias.ModuleID = NewValue
    pResFileAlias.ModuleID = NewValue
End Property


Property Get Name() As String
    Name = pAlias.Name
End Property

Property Let Name(NewValue As String)
    pAlias.Name = NewValue
End Property


Property Get DataType() As String
    DataType = pAlias.DataType
End Property

Property Let DataType(NewValue As String)
    pAlias.DataType = NewValue
End Property


Property Get Alias() As VBAlias
    Set Alias = pAlias
End Property


Property Get FilePath() As String
    FilePath = pFilePath
End Property

Property Let FilePath(NewValue As String)
    pFilePath = NewValue
    FileName = GetFileName(NewValue)
End Property


Property Get FileName() As String
    FileName = pFileAlias.Name
End Property

Property Let FileName(NewValue As String)
    pFileAlias.Name = NewValue
End Property


Property Get FileAlias() As VBAlias
    Set FileAlias = pFileAlias
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

Property Get LineCount() As Long
    LineCount = pLineCount
End Property

Property Get Lines(Index As Long) As String
    Lines = pLines(Index)
End Property

Property Let Lines(Index As Long, NewValue As String)
    pLines(Index) = NewValue
End Property

Property Get LineProtected(Index As Long) As Boolean
    LineProtected = pLineProtected(Index)
End Property

Property Let LineProtected(Index As Long, NewValue As Boolean)
    pLineProtected(Index) = NewValue
End Property

Property Get FileData() As String
    FileData = pFileData
End Property

Property Let FileData(NewValue As String)
    pFileData = NewValue
End Property


Property Get Code() As String
    Code = pCode
End Property

Property Let Code(NewValue As String)
    pCode = NewValue
End Property


Property Get Encrypt() As Boolean
    Encrypt = pEncrypt
End Property

Property Let Encrypt(NewValue As Boolean)
    pEncrypt = NewValue
End Property


Property Get BlockCount() As Long
    BlockCount = pBlockCount
End Property

Property Let BlockCount(NewValue As Long)
    ReDim Preserve pBlocks(NewValue)
    pBlockCount = NewValue
End Property

Property Get Blocks(Index As VBModuleCodeBlocks) As VBCodeBlock
    Set Blocks = pBlocks(Index)
End Property

Property Set Blocks(Index As VBModuleCodeBlocks, NewObject As VBCodeBlock)
    Set pBlocks(Index) = NewObject
End Property


Property Get PropertyCount() As Long
    PropertyCount = pPropertyCount
End Property

Property Let PropertyCount(NewValue As Long)
    ReDim Preserve pProperties(NewValue)
    pPropertyCount = NewValue
End Property

Property Get Properties(Index As Long) As VBProperty
    Set Properties = pProperties(Index)
End Property

Property Set Properties(Index As Long, NewObject As VBProperty)
    Set pProperties(Index) = NewObject
End Property


Property Get ControlCount() As Long
    ControlCount = pControlCount
End Property

Property Let ControlCount(NewValue As Long)
    ReDim Preserve pControls(NewValue)
    pControlCount = NewValue
End Property

Property Get Controls(Index As Long) As VBControl
    Set Controls = pControls(Index)
End Property

Property Set Controls(Index As Long, NewObject As VBControl)
    Set pControls(Index) = NewObject
End Property


Property Get DeclarationCount() As Long
    DeclarationCount = pDeclarationCount
End Property

Property Let DeclarationCount(NewValue As Long)
    ReDim Preserve pDeclarations(NewValue)
    pDeclarationCount = NewValue
End Property


Property Get Declarations(Index As Long) As VBDeclaration
    Set Declarations = pDeclarations(Index)
End Property

Property Set Declarations(Index As Long, NewObject As VBDeclaration)
    Set pDeclarations(Index) = NewObject
End Property


Property Get EventCount() As Long
    EventCount = pEventCount
End Property

Property Let EventCount(NewValue As Long)
    ReDim Preserve pEvents(NewValue)
    pEventCount = NewValue
End Property


Property Get Events(Index As Long) As VBEvent
    Set Events = pEvents(Index)
End Property

Property Set Events(Index As Long, NewObject As VBEvent)
    Set pEvents(Index) = NewObject
End Property


Property Get EnumCount() As Long
    EnumCount = pEnumCount
End Property

Property Let EnumCount(NewValue As Long)
    ReDim Preserve pEnums(NewValue)
    pEnumCount = NewValue
End Property


Property Get Enums(Index As Long) As VBEnum
    Set Enums = pEnums(Index)
End Property

Property Set Enums(Index As Long, NewObject As VBEnum)
    Set pEnums(Index) = NewObject
End Property


Property Get TypeCount() As Long
    TypeCount = pTypeCount
End Property

Property Let TypeCount(NewValue As Long)
    ReDim Preserve pTypes(NewValue)
    pTypeCount = NewValue
End Property


Property Get Types(Index As Long) As VBType
    Set Types = pTypes(Index)
End Property

Property Set Types(Index As Long, NewObject As VBType)
    Set pTypes(Index) = NewObject
End Property


Property Get VariableCount() As Long
    VariableCount = pVariableCount
End Property

Property Let VariableCount(NewValue As Long)
    ReDim Preserve pVariables(NewValue)
    pVariableCount = NewValue
End Property

Property Get Variables(Index As Long) As VBVariable
    Set Variables = pVariables(Index)
End Property

Property Set Variables(Index As Long, NewObject As VBVariable)
    Set pVariables(Index) = NewObject
End Property


Property Get SubCount() As Long
    SubCount = pSubCount
End Property

Property Let SubCount(NewValue As Long)
    ReDim Preserve pSubs(NewValue)
    pSubCount = NewValue
End Property

Property Get Subs(Index As Long) As VBSub
    Set Subs = pSubs(Index)
End Property

Property Set Subs(Index As Long, NewObject As VBSub)
    Set pSubs(Index) = NewObject
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

Property Set Strings(Index As Long, NewObj As VBString)
    Set pStrings(Index) = NewObj
End Property


Function GetModuleName(ModCode As String) As String

    Dim Lines() As String
    Dim c As String
    Dim i As Long
    Dim ValueName As String
    Dim Value As String
    Dim s As String
    
    Lines = Split(ModCode, vbNewLine)

    For i = 0 To UBound(Lines)

        If Left(Lines(i), Len("Attribute VB_Name = ")) = "Attribute VB_Name = " Then
            GetModuleName = FilterString(Mid(Lines(i), Len("Attribute VB_Name = ") + 1), Chr(34))
            Exit Function
        End If
        
    Next i

End Function


Function LoadModuleFile(ModuleFilePath As String)
    
    Dim ThisModule As VBModule
    Dim Encrypt As Boolean
    
    Dim i As Long
    Dim ControlCount As Long
    Dim ControlNames() As String
    
    Dim Vars() As VBAlias
    
    
    FilePath = ModuleFilePath
    
    FileData = ReadFile(ModuleFilePath)
    
    
    Code = AdaptModule(FileData)
    
    
    pLines = Split(Code, vbNewLine)
    pLineCount = UBound(pLines) + 1
    ReDim pLineProtected(UBound(pLines))
    
    FindBlocks
    
    Name = GetModuleName(Blocks(C_Attributes).Code)
    
'    If Name = "MyRegistryClass" Then
'        Debug.Assert False
'    End If
    
    Select Case DataType
        Case "Module", "Class"
            
        Case "Form", "UserControl", "PropertyPage", "Designer"
            
            LoadProperties
            
            'ResFilePath = ParentProject.BaseDirPath & GetResFileName(DataType, Blocks(B_Properties).Code)
            ResFilePath = ParentProject.BaseDirPath & GetResFileName(DataType, Me.Code)
            If ResFileName <> "" Then
                ResFileAlias.Name = ResFileName
            End If
            Encrypt = True
            
        Case Else
            
            Debug.Print DataType, ModuleFilePath
            Debug.Assert False
            Exit Function
            
    End Select
    
    Select Case DataType
            
        Case "Form", "UserControl", "PropertyPage"
            
            FindControls
            
        Case Else
                
    End Select
    
    
    'FindDeclarations
    'FindEvents
    'FindEnums
    'FindTypes
    FindVariables
    FindSubs
    FindStrings
    
    If pCaptionCount > 0 Then

        AddCaptionInitCode
        
        Code = AdaptModule(FileData)
    
        pLines = Split(Code, vbNewLine)
        pLineCount = UBound(pLines) + 1
        ReDim pLineProtected(UBound(pLines))
        
        FindBlocks
        
        FindSubs
        FindStrings
        
    End If
    
    
    FindSubStart
    
    
    'If pCaptionCount > 0 Then
    '    Debug.Print GetCaptionsLoaderSubString
    'End If
    
    'Debug.Print Me.ID, Me.Name
    
End Function

Private Function GetResFileName(ModuleDataType As String, Code As String) As String
    
    Const ResFileExtensions As String = ".frx|.ctx|.pgx|.dsx"
    
    Dim ResFileExt As String
    Dim s As String
    Dim x As Long
    Dim ResLine As String
    Dim StartPos As Long
    Dim EndPos As Long
    
    ResFileExt = DataTypeResExt(DataType)
    
    s = ResFileExt & Chr(34) & ":"
    
    x = InStr(1, Code, s)
    
    If x = 0 Then
        Exit Function
    End If
    
    ResLine = GetLine(Code, x)
    x = InStr(1, ResLine, s)
    
    StartPos = InStrRev(ResLine, Chr(34), x) + 1
    EndPos = x + Len(ResFileExt)
    
    GetResFileName = Mid(ResLine, StartPos, EndPos - StartPos)
    
End Function

Function DataTypeResExt(DataType As String) As String
    Select Case DataType
            Case "Form"
                DataTypeResExt = ".frx"
            Case "UserControl"
                DataTypeResExt = ".ctx"
            Case "PropertyPage"
                DataTypeResExt = ".pgx"
            Case "Designer"
                DataTypeResExt = ".dsx"
        End Select
End Function

Function DataTypeExt(DataType) As String
    Select Case DataType
        Case "Module"
            DataTypeExt = ".bas"
        Case "Class"
            DataTypeExt = ".cls"
        Case "Form"
            DataTypeExt = ".frm"
        Case "UserControl"
            DataTypeExt = ".ctl"
        Case "PropertyPage"
            DataTypeExt = ".pag"
        Case "Designer"
            DataTypeExt = ".dsr"
        Case "ResFile32"
            DataTypeExt = ".res"
    End Select
End Function


Private Function FindBlocks()
    
    Const HeadLineIndicator             As String = "VERSION"
    Const BeginIndicator                As String = "Begin"
    Const EndIndicator                  As String = "End"
    Const AttributeIndicator            As String = "Attribute"
    
    Const SubIndicatorsString           As String = "Sub |Function |Property |Private Sub |Private Function |Private Property |Public Sub |Public Function |Public Property "
    Const ScopeIndicatorsString         As String = "Private|Public|Global"
    Const DeclarationIndicatorsString   As String = "Declare|Type|Enum"
    
    Dim i As Long
    Dim l As Long
    Dim r As Long
    
    'Dim Lines() As String
    Dim Line As String
    
    Dim SubIndicators() As String
    Dim DeclarationIndicators() As String
    
    Dim FirstLine As Long
    Dim LastLine As Long
    
    Dim Found As Boolean
    
    Dim HeadLine As String
    Dim DefBlock As String
    Dim AttrBlock As String
    Dim DeclBlock As String
    Dim SubBlock As String
    
    BlockCount = 5
    For i = 1 To 5
        Set Blocks(i) = New VBCodeBlock
        Set Blocks(i).ParentModule = Me
    Next i
        
    SubIndicators() = Split(SubIndicatorsString, "|")
    DeclarationIndicators() = Split(DeclarationIndicatorsString, "|")
        
    'pLines() = Split(Code, vbNewLine)
        
    If DataType <> "Module" Then
        
        HeadLine = pLines(0)
        Blocks(1).Code = HeadLine
        Blocks(1).FirstLine = 0
        Blocks(1).LastLine = 0
        Blocks(1).LineCount = 1
        'Debug.Print HeadLine
        
        FirstLine = 0
        For i = 1 To UBound(pLines)
            
            If UCase(Left(pLines(i), Len(BeginIndicator))) = UCase(BeginIndicator) Then
                FirstLine = i
                Exit For
            End If
            
        Next i
        
        If UCase(Left(pLines(FirstLine), Len(BeginIndicator))) = UCase(BeginIndicator) Then
            
                    
            For i = FirstLine To UBound(pLines)
                    
                If UCase(Left(pLines(i), Len(EndIndicator))) = UCase(EndIndicator) Then
                    LastLine = i
                    DefBlock = JoinLines(pLines(), FirstLine, LastLine)
                    Blocks(2).Code = DefBlock
                    Blocks(2).FirstLine = FirstLine
                    Blocks(2).LastLine = LastLine
                    Blocks(2).LineCount = LastLine - FirstLine
                    'Debug.Print Quote(DefBlock)
                    Exit For
                End If
                    
            Next i
            
            LastLine = LastLine + 1
        End If
        
    End If
    
        
    FirstLine = LastLine
    For i = FirstLine To UBound(pLines)
        
        If UCase(Left(pLines(i), Len(AttributeIndicator))) = UCase(AttributeIndicator) Then
            FirstLine = i
            LastLine = 0
            For l = FirstLine To UBound(pLines)
                If UCase(Left(pLines(l), Len(AttributeIndicator))) <> UCase(AttributeIndicator) Then
                    Found = True
                    LastLine = l - 1
                    Exit For
                End If
            Next l
            If Not Found Then
                If LastLine = 0 Then
                    LastLine = UBound(pLines)
                End If
            End If
            AttrBlock = JoinLines(pLines(), FirstLine, LastLine)
            Blocks(3).Code = AttrBlock
            Blocks(3).FirstLine = FirstLine
            Blocks(3).LastLine = LastLine
            Blocks(3).LineCount = LastLine - FirstLine
            'Debug.Print Quote(AttrBlock)
            Exit For
        End If
    Next i
        
        
    FirstLine = LastLine + 1
    For i = FirstLine To UBound(pLines)
        'If AnyInText(pLines(i), SubIndicators()) <> 0 Then
         If LineBeginIsIn(pLines(i), SubIndicators) <> -1 Then
            If AnyInText(pLines(i), DeclarationIndicators) = 0 Then
                
                LastLine = i - 1
                DeclBlock = JoinLines(pLines(), FirstLine, LastLine)
                Blocks(4).Code = DeclBlock
                Blocks(4).FirstLine = FirstLine
                Blocks(4).LastLine = LastLine
                Blocks(4).LineCount = LastLine - FirstLine
                'Debug.Print Quote(DeclBlock)
                GoTo Continue
                
            End If
        End If
    Next i
    
    LastLine = UBound(pLines)
    DeclBlock = JoinLines(pLines(), FirstLine, LastLine)
    Blocks(4).Code = DeclBlock
    Blocks(4).FirstLine = FirstLine
    Blocks(4).LastLine = LastLine
    Blocks(4).LineCount = LastLine - FirstLine
    'Debug.Print Quote(DeclBlock)
    
    Exit Function
       
Continue:

    If DeclBlock <> "" Then
        FirstLine = LastLine + 1
    End If
        
    LastLine = UBound(pLines)
        
    SubBlock = JoinLines(pLines(), FirstLine, LastLine)
    'Clipboard.SetText SubBlock
    'Debug.Print Quote(SubBlock)
        
    Blocks(5).Code = SubBlock
    Blocks(5).FirstLine = FirstLine
    Blocks(5).LastLine = LastLine
    Blocks(5).LineCount = LastLine - FirstLine
    
'    For i = 1 To 5
'
'        Debug.Print JoinLines(pLines, Blocks(i).FirstLine, Blocks(i).LastLine)
'
'    Next i
    
End Function

Private Sub FindSubStart()
    
    Const SubIndicatorsString           As String = "Sub |Function |Property |Private Sub |Private Function |Private Property |Public Sub |Public Function |Public Property "
    
    Dim SubIndicators() As String
    
    Dim i As Long
    Dim r As Long
    Dim StartPos As Long
    Dim sLine As String
    
    SubIndicators = Split(SubIndicatorsString, "|")
    
    StartPos = Len(FileData)
    
    For i = LBound(SubIndicators) To UBound(SubIndicators)
        
        r = InStr(1, FileData, SubIndicators(i), vbTextCompare)
        
        If r > 0 Then
            
            sLine = GetLine(FileData, r)
            
            If (InStr(1, sLine, "(") > 0) And (InStr(1, sLine, ")") > 0) Then
        
                If r < StartPos Then
                    
                    StartPos = r
                
                End If
            
            End If
            
        End If
        
    Next i
    
    pSubStart = StartPos
    
End Sub


Private Function FindControls()

    'Const FormBeginIndicator As String = "Begin VB.Form"
    'Const FormEndIndicator As String = "End"
    Const ControlBeginIndicator As String = "Begin"
    Const ControlEndIndicator As String = "End"
    Const ProtectedControlsIndicator  As String = " VB."
    
    Dim ControlBeginKeyWord As String
    Dim ControlEndKeyWord As String
    
    Dim l As Long
    Dim i As Long
    Dim r As Long
    Dim x As Long
    
    Dim FirstLine As Long
    Dim LastLine As Long
    
    Dim ShiftLevel As Long
    Dim SpaceCount As Long
    
    Dim FoundSubLevel As Boolean
    
    Dim StartPos As Long
    Dim EndPos As Long
    Dim CodeStr As String
    Dim CtlDef As String
    
    Dim CtlName As String
    Dim TypeStr As String
    Dim LineStr As String
    Dim ControlNames() As String
    Dim Lines() As String
    Dim Words() As String
    
    Dim ChildBlock As VBCodeBlock
    Dim Control As VBControl
    
    ReDim ControlNames(0)
    
    
    
Begin:
    
    ShiftLevel = ShiftLevel + 1
    SpaceCount = ShiftLevel * 3
    FoundSubLevel = False
    
    ControlBeginKeyWord = String(SpaceCount, " ") & ControlBeginIndicator & " "
    ControlEndKeyWord = String(SpaceCount, " ") & ControlEndIndicator
    
    l = Blocks(B_Properties).FirstLine
    Do
    
        If Left(pLines(l), Len(ControlBeginKeyWord)) = ControlBeginKeyWord Then
            
            FirstLine = l
            
            For i = l To Blocks(B_Properties).LastLine
                
                If Not FoundSubLevel Then
                    If InStr(1, pLines(i), String(3, " ") & ControlBeginKeyWord) <> 0 Then
                        FoundSubLevel = True
                    End If
                End If
                
                If Left(pLines(i), Len(ControlEndKeyWord)) = ControlEndKeyWord Then
                    
                    LastLine = i
                    
                    ControlCount = ControlCount + 1
                    
                    ReDim Preserve ControlNames(ControlCount)
                    ControlNames(ControlCount) = CtlName
                    
                    Set Controls(ControlCount) = New VBControl
                    Set Control = Controls(ControlCount)
                    Set Control.ParentModule = Me
                    
                    Control.ID = ControlCount
                    
                    Control.Block.FirstLine = FirstLine
                    Control.Block.LastLine = LastLine
                    
                    Control.Alias.AliasType = VBVariable + VBControl + VBWithEvents
                    Control.Alias.Scope = VBLocal + VBPrivate + VBPublic
                    
                    Control.LoadControlCode FirstLine, LastLine
                    
                    l = LastLine
                    Exit For
                    
                End If
            Next i
            
        End If
        
         l = l + 1
    Loop While l <= Blocks(B_Properties).LastLine
    
    If FoundSubLevel Then
        GoTo Begin
    End If
    
    Exit Function
    
    'CodeStr = Blocks(B_Properties).Code
    '
    'x = 1
    'Do
    '    StartPos = InText(x, CodeStr, ControlBeginIndicator)
    '    If StartPos <> 0 Then
    '        x = StartPos + Len(ControlBeginIndicator)
    '        EndPos = InText(x, CodeStr, ControlEndIndicator)
    '        If EndPos <> 0 Then
    '
    '            EndPos = EndPos + Len(ControlEndIndicator)
    '            CtlDef = Mid(CodeStr, StartPos, EndPos - StartPos)
    '
    '
    '            Lines = Split(CtlDef, vbNewLine)
    '            LineStr = RTrim(LTrim(Lines(0)))
    '            Words = Split(LineStr, " ")
    '
    '            CtlName = Words(2)
    '
    '            If InStrArray(CtlName, ControlNames()) = -1 Then
    '
    '
    '                ControlCount = ControlCount + 1
    '
    '                ReDim Preserve ControlNames(ControlCount)
    '                ControlNames(ControlCount) = CtlName
    '
    '                Set Controls(ControlCount) = New VBControl
    '
    '                Set Control = Controls(ControlCount)
    '                Set Control.ParentModule = Me
    '
    '                Control.ID = ControlCount
    '                'Control.Name = CtlName
    '                'Control.DataType = TypeStr
    '                Control.Alias.AliasType = VBVariable + VBControl + VBWithEvents
    '                Control.Alias.Scope = VBLocal + VBPrivate + VBPublic
    '
    '                Control.LoadControlCode CtlDef
    '
    '                'Control.Block.StartPos = StartPos
    '                'Control.Block.Length = (EndPos - StartPos)
    '                'Control.Block.EndPos = EndPos
    '
    '                    'Debug.Print "'" & Control.CtlName & "'"
    '
    '            End If
    '
    '        End If
    '    End If
    'Loop While StartPos <> 0
    
        'Debug.Print "'" & CodeStr & "'"
    
        'FindControls = .ControlCount

End Function

Private Function FindVariables()
    
    Const AntiIndicatorsString      As String = "Declare|Type|Enum|Event"
    Const DeclIndicatorsString      As String = "Dim|Private|Public|Global"
    
    Const LongDataTypeIndicator     As String = "As"
    
    Const TypeSignsString           As String = "@|#|%|&|!|$"
    Const DataTypesString           As String = "Boolean|Byte|Currency|Date|Double|Integer|Long|Object|Single|String|Variant"
    
    Const ConstIndicator            As String = "Const"
    
    
    Dim AntiIndicators()            As String
    Dim DeclIndicators()            As String
    
    Dim DataTypes()                 As String
    Dim TypeSigns()                 As String
    
    Dim VarsStr                     As String
    Dim Vars()                      As String
    Dim VarLine                     As String

    
    Dim l As Long
    Dim i As Long
    Dim n As Long
    Dim y As Long
    
    'Dim a As Long
    Dim r As Long
    Dim x As Long
    Dim z As Long
    Dim c As Long
    
    Dim Found As Boolean
    Dim IsConstant As Boolean
    
    
    If Trim(Code) = "" Then
        Exit Function
    End If
    
    
    AntiIndicators() = Split(AntiIndicatorsString, "|")
    DeclIndicators() = Split(DeclIndicatorsString, "|")
    
    DataTypes() = Split(DataTypesString, "|")
    TypeSigns() = Split(TypeSignsString, "|")
    
    'Debug.Print Code
    
'    If Me.Name = "VBAlias" Then
'
'        Debug.Assert False
'
'    End If
    
    'Debug.Print JoinLines(pLines, Blocks(D_Declinations).FirstLine, Blocks(D_Declinations).LastLine)
    
    For l = Blocks(D_Declinations).FirstLine To Blocks(D_Declinations).LastLine
        
        If (AnyInText(Lines(l), AntiIndicators()) = 0) And (Trim(Lines(l)) <> "") Then
            
            z = AnyInText(Lines(l), DeclIndicators())
            If z <> 0 Then
                
                r = InText(1, Lines(l), DeclIndicators(z - 1))
                
                'If InText(1, Lines(l), LongDataTypeIndicator) <> 0 Then
                '    Found = True
                'Else
                '    a = AnyInText(Lines(l), TypeSigns, False)
                '    Found = (a <> 0)
                'End If
                
                'If Found Then
                    
                    
                    VarsStr = Mid(Lines(l), r + Len(DeclIndicators(z - 1)) + 1)
                    VarsStr = LTrim(VarsStr)
                    x = InText(1, VarsStr, ConstIndicator)
                    If x <> 0 Then
                        IsConstant = True
                        VarsStr = Mid(VarsStr, Len(ConstIndicator) + 1)
                        VarsStr = LTrim(VarsStr)
                    Else
                        IsConstant = False
                    End If
                    Vars() = Split(VarsStr, ", ")
                    
                
                    For n = 0 To UBound(Vars)
                    
                        VarLine = DeclIndicators(z - 1) & " "
                        If IsConstant Then
                            VarLine = VarLine & "Const "
                        End If
                        VarLine = VarLine & Vars(n)
                    
                        VariableCount = VariableCount + 1
                        Set Variables(VariableCount) = New VBVariable
                    
                        With Variables(VariableCount)
                        
                            Set .ParentModule = Me
                            .ID = VariableCount
                            .LoadLine VarLine
                            
                            'Debug.Print .Name
                            
                        End With
                        
                    
                    Next n
                    
                'End If
            
            End If
            
        End If
            
    Next l
    
    
End Function


Private Function FindSubs0()
    
    Const SubIndicators As String = "Sub|Function|Property"
        
    'Dim Lines() As String
    Dim Indicators() As String
    
    Dim i As Long
    Dim l As Long
    Dim y As Long
    Dim z As Long
    Dim r As Long
    
    Dim FirstLine As Long
    Dim LastLine As Long
    Dim SubCode As String
    
    If Blocks(E_Subs).Code = "" Then
        Exit Function
    End If
        
    'Lines = Split(Blocks(E_Subs).Code, vbNewLine)
    Indicators = Split(SubIndicators, "|")
        
    
    l = Blocks(E_Subs).FirstLine
    Do
        
        i = 0
        Do
        
            r = InText(1, Lines(l), Indicators(i))
            
            If r <> 0 Then
                
                    
                FirstLine = l
                
                y = l + 1
                Do
                    
                    r = InText(1, Lines(y), "End " & Indicators(i))
                    If r <> 0 Then
                        LastLine = y
                        Exit Do
                    End If
                    
                    y = y + 1
                    
                Loop While y <= Blocks(E_Subs).LastLine
                    
                
                SubCount = SubCount + 1
                
                Set Subs(SubCount) = New VBSub
                Set Subs(SubCount).ParentModule = Me
                
                Subs(SubCount).ID = SubCount
                Subs(SubCount).Block.FirstLine = FirstLine
                Subs(SubCount).Block.LastLine = LastLine
                
                'Debug.Print Me.Name
                
                'Subs(SubCount).LoadSubCode FirstLine, LastLine
                
                l = LastLine
                GoTo Continue
                
                
            End If
            
            i = i + 1
        Loop While i <= UBound(Indicators)
        
Continue:
        
        DoEvents
        l = l + 1
    Loop While l <= Blocks(E_Subs).LastLine
    
    For i = 1 To SubCount
        Subs(i).LoadSubCode Subs(i).Block.FirstLine, Subs(i).Block.LastLine
        RaiseEvent FileProgress(100 / SubCount * i)
    Next
    
End Function

Private Function FindSubs()
    
    Const SubIndicators As String = "Sub|Function|Property"
    
    Dim Indicators() As String
    
    Dim EndIndicator As String
    
    Dim l As Long
    Dim i As Long
    Dim r As Long
    
    Dim FirstLine As Long
    Dim LastLine As Long
    
    Indicators = Split(SubIndicators, "|")
    
    SubCount = 0
        
    'Debug.Print Me.Name
        
    For l = Blocks(E_Subs).FirstLine To Blocks(E_Subs).LastLine
        
        If FirstLine = 0 Then
            
            For i = 0 To UBound(Indicators)
                
                r = InText(1, pLines(l), Indicators(i))
                
                If IsQuoted(pLines(l), r, Len(Indicators(i))) Then
                    r = 0
                End If
                
                If r <> 0 Then
                    FirstLine = l
                    Exit For
                End If
                
            Next i
            
        End If
        
        If FirstLine <> 0 Then
            
              
            For i = 0 To UBound(Indicators)
                
                EndIndicator = "End " & Indicators(i)
                
                r = InText(1, Lines(l), EndIndicator)
                
                If IsQuoted(pLines(l), r, Len(EndIndicator)) Then
                    r = 0
                End If
                
                If r <> 0 Then
                    LastLine = l
                    
                    SubCount = SubCount + 1
                
                    Set Subs(SubCount) = New VBSub
                    Set Subs(SubCount).ParentModule = Me
                    
                    Subs(SubCount).ID = SubCount
                    Subs(SubCount).Block.FirstLine = FirstLine
                    Subs(SubCount).Block.LastLine = LastLine
                    
                    'Debug.Assert (SubCount <> 97)
                    'Debug.Print JoinLines(pLines, FirstLine, LastLine)
                    
                    FirstLine = 0
                    LastLine = 0
                    Exit For
                End If
                
            Next i
            
            
        End If
        
    Next l
    
    
    
    For i = 1 To SubCount
        Subs(i).LoadSubCode Subs(i).Block.FirstLine, Subs(i).Block.LastLine
        RaiseEvent FileProgress(100 / SubCount * i)
    Next
    
End Function

Function PreProcess()
    
    Dim i As Long
    
     'Add Caption Loader Sub Code
   
        

    
    For i = 1 To StringCount
        
        If Strings(i).Checked Then
        
            Strings(i).ReplaceString
        
        End If
        
        RaiseEvent FileProgress(100 / StringCount * i)
        If Cancelled Then Exit Function
        
    Next
    
    
End Function

Function CryptModule()
    
    Dim i As Long
    Dim c As Long
    
    Dim ControlNames() As String
    Dim ControlNameCount As Long
    Dim OldGlobalName As String
    Dim NewGlobalName As String
    
    ReDim ControlNames(0)
    
    For i = 1 To PropertyCount
        
        Properties(i).Crypt
    
    Next i
    
      'Crypt Controls
    For i = 1 To ControlCount
        
        Controls(i).CryptControl
        
        If InStrArray(Controls(i).Name, ControlNames) = -1 Then
            
            ControlNameCount = ControlNameCount + 1
            ReDim Preserve ControlNames(ControlNameCount)
            ControlNames(ControlNameCount) = Controls(i).Name
            
            OldGlobalName = Me.Name & "." & Controls(i).Alias.Name
            NewGlobalName = Me.Alias.NewName & "." & Controls(i).Alias.NewName
            
            ReplaceInSubs Controls(i).Alias.Name, Controls(i).Alias.NewName
            ParentProject.GlobalReplace OldGlobalName, NewGlobalName

        End If
        
    Next i
    
    ' Crypt Subs
    For i = 1 To SubCount
        
        Subs(i).CryptSub
        
        RaiseEvent FileProgress(100 / SubCount * i)
        If Cancelled Then Exit Function
    
    Next i
    
    
    For i = 1 To DeclarationCount
        
        Declarations(i).Crypt
    Next i
    
    For i = 1 To EventCount
        Events(i).Crypt
    Next i
    
    For i = 1 To VariableCount
        
'        If Variables(i).Alias.Name = "AliasType" Then
'            Debug.Assert False
'        End If
        
        ReplaceInMod Variables(i).Alias.Name, Variables(i).Alias.NewName
        
        If (Variables(i).Alias.Scope And VBPublic) Then
            ParentProject.GlobalReplace Variables(i).Alias.Name, Variables(i).Alias.NewName
        End If
    
    Next i
    
    'Crypt Module
    If ResFileName <> "" Then
        
        
        
        ReplaceInLines pLines, 0, Me.LineCount - 1, Quote(ResFileAlias.Name), Quote(ResFileAlias.NewName)
        ReplaceInLines pLines, 0, Me.LineCount - 1, ResFileAlias.Name, ResFileAlias.NewName
        
        'ReplaceInLines pLines, Blocks(B_Properties).FirstLine, Blocks(B_Properties).LastLine, Quote(ResFileAlias.Name), Quote(ResFileAlias.NewName)
        'ReplaceInLines pLines, Blocks(B_Properties).FirstLine, Blocks(B_Properties).LastLine, ResFileAlias.Name, ResFileAlias.NewName
    
    End If
    
    For i = 1 To ParentProject.ModuleCount
        With ParentProject.Modules(i)
            
            .ReplaceInMod Quote(Alias.Name), Quote(Alias.NewName)
            .ReplaceInMod Alias.Name, Alias.NewName
            
        End With
    Next i
    
    For i = 1 To TypeCount
        Types(i).Crypt
    Next i
    
    For i = 1 To EnumCount
        Enums(i).Crypt
    Next i
    
    
        
    If My.Config.EncryptStrings Then
        
        For i = 1 To StringCount
            Strings(i).EncryptString
        Next i
        
    End If
        

    
    RaiseEvent FileProgress(100)
    
End Function

Function LoadProperties()
    
    Const BeginKeyWord As String = "Begin"
    Const EndKeyWord1 As String = "   Begin"
    Const EndKeyWord2 As String = "End"
    
    Dim l As Long
    Dim i As Long
    Dim y As Long
    Dim Lines() As String
    Dim LineStr As String
    Dim PropStr As String
    Dim FirstLine As Long
    Dim LastLine As Long
    Dim FoundEnd As Boolean
    
    
    For i = Blocks(B_Properties).FirstLine To Blocks(B_Properties).LastLine
        
        If Left(pLines(i), Len(BeginKeyWord)) = BeginKeyWord Then
            FirstLine = i + 1
            For y = FirstLine To Blocks(B_Properties).LastLine
                If Left(pLines(y), Len(EndKeyWord1)) = EndKeyWord1 Then
                    FoundEnd = True
                ElseIf Left(pLines(y), Len(EndKeyWord2)) = EndKeyWord2 Then
                    FoundEnd = True
                End If
                If FoundEnd Then
                    LastLine = y - 1
                    GoTo Continue:
                End If
            Next y
        End If
        
    Next i
    
    Exit Function
    
Continue:

    'Debug.Print JoinLines(pLines, FirstLine, LastLine)
    
    For l = FirstLine To LastLine
        
        'Debug.Print Lines(c)
        
        PropertyCount = PropertyCount + 1
        Set Properties(PropertyCount) = New VBProperty
        Set Properties(PropertyCount).ParentModule = Me
        
        With Properties(PropertyCount)
            
            If InText(1, pLines(l), "EndProperty") = 0 Then
                .LoadLine l
                'Debug.Print .Name
            End If
        
            If InText(1, pLines(l), "BeginProperty") <> 0 Then
            
                For i = l + 1 To LastLine
                
                    If InText(1, Lines(i), "EndProperty") = 0 Then
                        
                        .ChildCount = .ChildCount + 1
                        Set .Childs(.ChildCount) = New VBProperty
                        Set .Childs(.ChildCount).ParentModule = Me
                        .Childs(.ChildCount).LoadLine i
                        
                    Else
                        
                        Exit For
                
                    End If
                    
                    
                Next i
            
            End If
            
        End With
        
    Next l
    
End Function


Function FindDeclarations()
    
    Const DeclIndicator = "Declare"
    
  
    Dim i As Long
    
    
    For i = Blocks(D_Declinations).FirstLine To Blocks(D_Declinations).LastLine
        
        If InText(1, pLines(i), DeclIndicator) <> 0 Then
            
            DeclarationCount = DeclarationCount + 1
            Set Declarations(DeclarationCount) = New VBDeclaration
            With Declarations(DeclarationCount)
                
                Set .ParentModule = Me
                .ID = DeclarationCount
                .LoadLine i
                
            End With
            
        End If
        
    Next i
    
End Function


Function FindEvents()
    
    Const EventIndicator = "Event"
    
  
    Dim i As Long
    
    
    For i = Blocks(D_Declinations).FirstLine To Blocks(D_Declinations).LastLine
        
        If InText(1, pLines(i), EventIndicator) <> 0 Then
            
            EventCount = EventCount + 1
            Set Events(EventCount) = New VBEvent
            With Events(EventCount)
                
                Set .ParentModule = Me
                .ID = EventCount
                .LoadLine i
                
            End With
            
        End If
        
    Next i
    
End Function


Function FindEnums()
    
    Const EnumIndicator As String = "Enum"
    Const EnumEndIndicator As String = "End Enum"
    
    Dim i As Long
    Dim l As Long
    Dim EnumDef As String
    Dim FirstLine As Long
    Dim LastLine As Long
    
    If InText(1, Blocks(D_Declinations).Code, EnumIndicator) = 0 Then
        Exit Function
    End If
    
    
    l = Blocks(D_Declinations).FirstLine
    Do
        
        If InText(1, pLines(l), EnumIndicator) <> 0 Then
            
            FirstLine = l
            
            For i = FirstLine To UBound(pLines)
                If InText(1, pLines(i), EnumEndIndicator) <> 0 Then
                    LastLine = i
                    
                    'EnumDef = JoinLines(pLines, FirstLine, LastLine)
                    
                    EnumCount = EnumCount + 1
                    Set Enums(EnumCount) = New VBEnum
                    With Enums(EnumCount)
                        Set .ParentModule = Me
                        .ID = EnumCount
                        .LoadEnumCode FirstLine, LastLine
                    End With
                    
                    l = LastLine
                    Exit For
                End If
            Next i
            
        End If
        
        l = l + 1
    Loop While l <= Blocks(D_Declinations).LastLine
    
End Function


Function FindTypes()
    
    Const TypeIndicator As String = "Type"
    Const TypeEndIndicator As String = "End Type"
    
    'Dim Lines() As String
    Dim i As Long
    Dim l As Long
    Dim TypeDef As String
    Dim FirstLine As Long
    Dim LastLine As Long
    
    'If InText(1, Blocks(D_Declinations).Code, TypeIndicator) = 0 Then
    '    Exit Function
    'End If
    
    'Lines() = Split(Blocks(D_Declinations).Code, vbNewLine)
    
    l = Blocks(D_Declinations).FirstLine
    Do
        
        If InText(1, Lines(l), TypeIndicator) <> 0 Then
            
            FirstLine = l
            
            For i = FirstLine To UBound(pLines)
                If InText(1, pLines(i), TypeEndIndicator) <> 0 Then
                    LastLine = i
                    
                    'TypeDef = JoinLines(Lines, FirstLine, LastLine)
                    
                    TypeCount = TypeCount + 1
                    Set Types(TypeCount) = New VBType
                    With Types(TypeCount)
                        .ID = TypeCount
                        Set .ParentModule = Me
                        .LoadTypeCode FirstLine, LastLine
                    End With
                    
                    l = LastLine
                    Exit For
                End If
            Next i
            
        End If
        
        l = l + 1
    Loop While l <= Blocks(D_Declinations).LastLine
    
End Function

Function AdaptModule(Code As String) As String
    
    Const KeyWord As String = "Attribute VB_"
    
    Dim Lines() As String
    Dim i As Long
    Dim l As Long
    Dim FirstLine As Long
    Dim LastLine As Long
    Dim Part1 As String
    Dim Part2 As String
    
    Lines = Split(Code, vbNewLine)
    
    For l = 0 To UBound(Lines)
        If Left(Lines(l), Len(KeyWord)) = KeyWord Then
            For i = l To UBound(Lines)
                If Left(Lines(i), Len(KeyWord)) <> KeyWord Then
                    FirstLine = i
                    LastLine = UBound(Lines)
                    GoTo Continue
                End If
            Next i
        End If
    Next l
    
    AdaptModule = Code
    Exit Function
    
    
Continue:
    
    Part1 = JoinLines(Lines, 0, FirstLine - 1)
    Part2 = JoinLines(Lines, FirstLine, LastLine)
    
    
    AdaptModule = Part1 & vbNewLine & AdaptCode(Part2)
       
End Function

Function ReplaceInMod(OldWord As String, NewWord As String)

    Dim i As Long
    
    For i = 0 To UBound(pLines)
        
        'If Not LineProtected(i) Then
            
            pLines(i) = ReplaceWords(pLines(i), OldWord, NewWord)
            
        'End If
        
    Next i

End Function

Function ReplaceInSubs(OldWord As String, NewWord As String)

    Dim i As Long
    
    For i = 1 To SubCount
        
        BlockReplace Subs(i).Block, OldWord, NewWord
        
    Next i

End Function

Function BlockReplace(Block As VBCodeBlock, OldWord As String, NewWord As String, Optional ReplaceInQuotes As Boolean = False)
    
    ReplaceInLines pLines, Block.FirstLine, Block.LastLine, OldWord, NewWord, ReplaceInQuotes
    
    'Debug.Print pLines(42)
    
End Function

Function ReplaceInCode(OldWord As String, NewWord As String)

    Dim i As Long
    
    For i = 1 To SubCount
        
        ReplaceInLines pLines, Subs(i).Block.FirstLine + 1, Subs(i).Block.LastLine - 1, OldWord, NewWord
        
    Next i

End Function

Function ReplaceStrings()
    
    Dim x As Long
    Dim y As Long
    
    For x = 1 To Me.SubCount
        
        For y = Me.Subs(x).Block.FirstLine To Me.Subs(x).Block.LastLine
            
            
            
        Next y
        
    Next x
    
End Function

Function Rebuild() As String
    
    Rebuild = Join(pLines, vbNewLine)
    
End Function


Property Get SubID(SubName As String) As Long

    Dim i As Long
    
    For i = 1 To SubCount
    
        If Subs(i).Name = SubName Then
            SubID = i
            Exit Property
        End If

    Next i
    
End Property


Function StringExists(Value As String) As Boolean

    Dim i As Long
    
    For i = 1 To pStringCount
        If pStrings(i).Value = Value Then
            StringExists = True
            Exit Function
        End If
    Next i

End Function

Function FindStrings()
    
    Const DebugKeyword As String = "Benutzername"
    
    Dim i As Long
    Dim l As Long
    Dim x As Long
    Dim n As Long
    Dim c As Long
    
    Dim bIgnore As Boolean
    Dim IgnoreIndicators() As String
    
    Dim q() As VBQuote
    
    
    StringCount = 0
    
    IgnoreIndicators = Split(fOptions.txtIgnore.Text, vbNewLine)
    
    For i = 1 To SubCount
        
        
        
        For l = Subs(i).Block.FirstLine To Subs(i).Block.LastLine
            
            'Debug.Print Subs(i).ParentModule.Lines(l)
            
            For n = 0 To UBound(IgnoreIndicators)
                
                If Trim(IgnoreIndicators(n)) <> "" Then
                
                    If LCase(Left(LTrim(pLines(l)), Len(IgnoreIndicators(n)))) = LCase(IgnoreIndicators(n)) Then
                        bIgnore = True
                        Exit For
                    Else
                        bIgnore = False
                    End If
                    
                End If
                
            Next n
            
            If Not bIgnore Then
            
                c = SplitQuotes(pLines(l), q)
        
                For x = 1 To c
                
                    If q(x).Quote <> "" Then
                    
                        If Not StringExists(q(x).Quote) Then
                            
                            StringCount = StringCount + 1
                            
                            Set pStrings(StringCount) = New VBString
                            Set pStrings(StringCount).ParentModule = Me
                            Set pStrings(StringCount).ParentSub = Subs(i)
                            
                            pStrings(StringCount).LineID = l
                            pStrings(StringCount).Index = x
                            pStrings(StringCount).LoadValue q(x).Quote
                            
                            If InStr(1, q(x).Quote, DebugKeyword, vbTextCompare) > 0 Then
                                
                            
                                
                                'Debug.Print q(x).Quote
                                'Debug.Assert False
                                
                            End If
                            
                            
                            
                            'Debug.Assert q(x).Quote <> "Beenden"
                            
                        End If
                        
                    End If
                    
                Next x
                
            End If
                
        Next l
        
    Next i
    
  
    
End Function

Function AddCaption(CaptionProperty As VBProperty) As Long
    
    pCaptionCount = pCaptionCount + 1
    ReDim Preserve pCaptions(pCaptionCount)
    
    Set pCaptions(pCaptionCount) = CaptionProperty
    
    AddCaption = pCaptionCount
    
End Function

Private Function GetCaptionsLoaderSubString() As String
    
    Dim i As Long
    Dim c As Long
    Dim x As Long
    Dim Value As String
    Dim CmdLines() As String
    Dim Lines() As String
    Dim SubCode As String
    Dim SpaceCount As Long
    Dim ControlProperty As String
    
    
    If pCaptionCount < 1 Then
        Exit Function
    End If
    
    ReDim CmdLines(1 To pCaptionCount)
    
    For i = 1 To pCaptionCount
        
        If Not pCaptions(i).ParentControl Is Nothing Then
            
            If pCaptions(i).ParentControl.Index <> "" Then
                'ControlProperty = String(4, " ") & "Me." & pCaptions(i).ParentControl.Name & "(" & pCaptions(i).ParentControl.Index & ")" & "." & pCaptions(i).Name
                ControlProperty = String(4, " ") & pCaptions(i).ParentControl.Name & "(" & pCaptions(i).ParentControl.Index & ")" & "." & pCaptions(i).Name
            Else
                'ControlProperty = String(4, " ") & "Me." & pCaptions(i).ParentControl.Name & "." & pCaptions(i).Name
                ControlProperty = String(4, " ") & pCaptions(i).ParentControl.Name & "." & pCaptions(i).Name
            End If
            
            Value = pCaptions(i).Value
            SpaceCount = Len(ControlProperty)
            
            Select Case pCaptions(i).Name
                
                Case "List"
                    
                    CmdLines(i) = vbNewLine & CmdLines(i)
                    
                    For x = 1 To pCaptions(i).ListCount
                        
                        'CmdLines(i) = "Me." & pCaptions(i).ParentControl.Name & "(" & CStr(x - 1) & ")" & "." & pCaptions(i).Name & " = " & Chr(34) & pCaptions(i).List(x) & Chr(34)
                        CmdLines(i) = CmdLines(i) & ControlProperty & "(" & CStr(x - 1) & ")" & " = " & Chr(34) & pCaptions(i).List(x) & Chr(34)
                        
                        If x < pCaptions(i).ListCount Then
                            
                            CmdLines(i) = CmdLines(i) & vbNewLine '& String(4, " ")
                            
                        End If
                        
                    Next x
                    
                    CmdLines(i) = CmdLines(i) & vbNewLine
                    
                Case Else
                    
                    CmdLines(i) = vbNewLine & CmdLines(i)
                    
                    CmdLines(i) = ControlProperty & " = "
                
                    If InStr(1, Value, vbNewLine) <> 0 Then
                        
                        Lines = Split(Value, vbNewLine)
                        
                        c = 0
                        For x = 0 To UBound(Lines())
                            
                            If Lines(x) <> "" Then
                                c = c + 1
                            End If
                            
                        Next x
                        
                        For x = 0 To c - 1

                            CmdLines(i) = CmdLines(i) & Chr(34) & Lines(x) & Chr(34)
                            
                            If (x + 1) < c Then
                                
                                CmdLines(i) = CmdLines(i) & " _ " & vbNewLine & String$(SpaceCount, " ") & " & "
                                
                            End If
                            
                        Next x
                        
                        CmdLines(i) = vbNewLine & CmdLines(i)
                        
                    Else
                        
                    
                        If Left$(Value, 1) = Chr(34) And Right$(Value, 1) = Chr(34) Then
                            
                            Value = Mid$(Value, 2, Len(Value) - 2)
                            
                        End If
                        
                        
                        CmdLines(i) = CmdLines(i) & Chr(34) & Value & Chr(34)
          
                    
                    End If
                    
                    If InStr(1, CmdLines(i), Chr(34)) = 0 Then
                        
                        Debug.Assert False
                        
                        Debug.Print Join(CmdLines, vbNewLine)
                        
                    End If

            End Select
            
            
            
        Else
        
            'CmdLines(i) = String(4, " ") & "Me." & pCaptions(i).Name & " = " & pCaptions(i).Value
            
            CmdLines(i) = String(4, " ") & pCaptions(i).Name & " = " & pCaptions(i).Value
            
        End If
        
    Next i
    
   
    pCaptionsLoaderSubName = "LoadControlCaptions"
    
    
    SubCode = "Private Sub " & pCaptionsLoaderSubName & "()" & vbNewLine & vbNewLine
    
    For i = 1 To pCaptionCount
        
        'SubCode = SubCode & String(4, " ") & CmdLines(i) & vbNewLine
        SubCode = SubCode & CmdLines(i) & vbNewLine
        
    Next i
    
    SubCode = SubCode & vbNewLine & "End Sub"
    
    GetCaptionsLoaderSubString = SubCode
    
End Function

Public Property Get InitSub() As VBSub
    
    Dim i As Long
    
    For i = 1 To SubCount
        
        If Right(Subs(i).Name, Len("_Load")) = "_Load" Then
            Set InitSub = Subs(i)
            Exit Property
        End If
    
    Next i
    
End Property

Private Sub AddCaptionInitCode()
    
    Dim SubCode As String
    Dim s As String
    
    Dim r As Long
    Dim StartPos As Long
    
    
    'Add Loader Sub Code
    'Lines(Subs(SubCount).Block.LastLine) = Lines(Subs(SubCount).Block.LastLine) & NewLine(1) & GetCaptionsLoaderSubString
    
    FileData = FileData & NewLine(3) & GetCaptionsLoaderSubString
    
    
    
    
    
    If Not InitSub Is Nothing Then
        
        r = InStr(1, FileData, "Sub " & InitSub.Name, vbTextCompare)
        
        If r <> 0 Then
            StartPos = InStr(r, FileData, ")")
            
            If StartPos <> 0 Then
                s = Left(FileData, StartPos) & NewLine(2) & String(4, " ") & pCaptionsLoaderSubName & NewLine(2) & Mid(FileData, StartPos + 1)
                FileData = s
            End If
        End If
        
    Else
        
        SubCode = "Private Sub " & Me.DataType & "_Load()" & vbNewLine & pCaptionsLoaderSubName & vbNewLine & "End Sub"
        FileData = FileData & NewLine(3) & SubCode
        
    End If
    
'    If Not InitSub Is Nothing Then
'
'        'Lines(InitSub.Block.FirstLine) = Lines(InitSub.Block.FirstLine) & vbNewLine & pCaptionsLoaderSubName
'
'    Else
'
'        'SubCode = "Private Sub " & Me.DataType & "_Load" & vbNewLine & pCaptionsLoaderSubName & vbNewLine & "End Sub"
'        'Lines(Subs(SubCount).Block.LastLine) = Lines(Subs(SubCount).Block.LastLine) & NewLine(2) & SubCode
'
'    End If
    
    
    
End Sub
'
'Property Get IsSpecialModule() As Boolean
'
'    Dim c As Long
'
'    For c = 1 To Me.ParentProject.SpecialModuleCount
'
'        If Me.ID = Me.ParentProject.SpecialModules(c).Index Then
'           IsSpecialModule = True
'           Exit For
'        End If
'
'    Next c
'
'End Property


Public Function GetStringIndex(StringValue As String) As Long
    
    Dim i As Long
    
    For i = 1 To StringCount
        
        If Strings(i).Value = StringValue Then
            
            GetStringIndex = i
            Exit Function
            
        End If
        
    Next i
    
End Function

Public Property Get SubStartPos() As Long
    
    SubStartPos = pSubStart

End Property
