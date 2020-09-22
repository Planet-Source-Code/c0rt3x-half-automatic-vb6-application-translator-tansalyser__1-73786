VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParentModule           As VBModule
Private pAlias                  As VBAlias
Private pCode                   As String

Private pHeadLine               As String
Private pBlock                  As VBCodeBlock
Private pSubType                As String

Private pParameterCount       As Long
Private pParameters()         As VBParameter

Private pVariableCount        As Long
Private pVariables()          As VBVariable

Private Sub Class_Initialize()
    
    Set pAlias = New VBAlias
    Set pBlock = New VBCodeBlock
    
    With pAlias
        .AliasType = VBSub
    End With
    
    ReDim pParameters(0)
    ReDim pVariables(0)
    
End Sub


Property Get ParentModule() As VBModule
    Set ParentModule = pParentModule
End Property

Property Set ParentModule(NewObject As VBModule)
    Set pParentModule = NewObject
    Set pAlias.ParentProject = pParentModule.ParentProject
    pAlias.ModuleID = NewObject.ID
End Property


Property Get Alias() As VBAlias
    Set Alias = pAlias
End Property


Property Get ID() As Long
    ID = pAlias.SubID
End Property

Property Let ID(NewValue As Long)
    pAlias.SubID = NewValue
    pAlias.ModuleID = ParentModule.ID
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


Property Get Code() As String
    Code = pCode
End Property

Property Let Code(NewValue As String)
    pCode = NewValue
End Property


Property Get HeadLine() As String
    HeadLine = Me.ParentModule.Lines(Me.Block.FirstLine)
End Property

Property Let HeadLine(NewValue As String)
    Debug.Assert False
    'pHeadLine = NewValue
End Property


Property Get Block() As VBCodeBlock
    Set Block = pBlock
End Property

Property Let Block(NewObject As VBCodeBlock)
    Set pBlock = NewObject
End Property


Property Get SubType() As String
    SubType = pSubType
End Property

Property Let SubType(NewValue As String)
    pSubType = NewValue
End Property


Property Get ParameterCount() As Long
    ParameterCount = pParameterCount
End Property

Property Let ParameterCount(NewValue As Long)
    ReDim Preserve pParameters(NewValue)
    pParameterCount = NewValue
End Property

Property Get Parameters(Index As Long) As VBParameter
    Set Parameters = pParameters(Index)
End Property

Property Set Parameters(Index As Long, NewObject As VBParameter)
    Set pParameters(Index) = NewObject
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


Sub LoadSubCode(FirstLine As Long, LastLine As Long)
        
    Dim sLine As String
    
    'Debug.Print SubCode
     
    'Code = SubCode
    'Block.Code = SubCode
    
    Block.FirstLine = FirstLine
    Block.LastLine = LastLine
    
    'HeadLine = ParentModule.Lines(FirstLine) 'GetSubDef(Code)
    Name = GetSubName(HeadLine)
    SubType = GetSubType(HeadLine)
    Select Case SubType
        Case "Function", "Property Get"
            DataType = GetSubDataType(HeadLine)
        Case Else
            
    End Select
    Alias.Scope = GetSubScope(HeadLine)
    
    FindVariables
    If Not IsEventHandler Then
        FindParameters
    End If
    
    If My.Config.EnableDebugMode Then
        
        If Me.ParentModule.Name <> "mDebug" Then
        
            sLine = "LogMsg CStr(Now) & "": " & Name & "()"""
            Me.ParentModule.Lines(FirstLine + 1) = sLine & vbNewLine & Me.ParentModule.Lines(FirstLine + 1)
        
        End If
        
    End If
    
    'If InStr(1, Me.Name, "Load", vbTextCompare) <> 0 Then
    '    Debug.Print Me.Name
    'End If
    
    'If Me.Name = "LoadConfig" Then
    '    Debug.Print Me.ParentModule.ID, Me.ParentModule.Name, Me.ID
    'End If
    
End Sub


Property Get IsEventHandler() As Boolean
        
    Const FixHandlersString As String = "Form_|MDIForm_|Class_|UserControl_|PropertyPage_|DataReport_"
    
    Dim i As Long
    Dim s As String
    Dim FixHandlers() As String
    
    If Name = "" Then Exit Property
    If InStr(1, Name, "_") = 0 Then Exit Property
    
    FixHandlers() = Split(FixHandlersString, "|")
    
    For i = 0 To UBound(FixHandlers)
        If Left(Name, Len(FixHandlers(i))) = FixHandlers(i) Then
            IsEventHandler = True
            Exit Property
        End If
    Next i
    
    For i = 1 To ParentModule.ControlCount
        If CBool((ParentModule.Controls(i).Alias.AliasType And VBWithEvents)) Then
            s = ParentModule.Controls(i).Name & "_"
            If Left(Name, Len(s)) = s Then
                IsEventHandler = True
                Exit Property
            End If
        End If
    Next i
    
    For i = 1 To ParentModule.VariableCount
        
        If ParentModule.Variables(i).HasEvents Then
            
            s = ParentModule.Variables(i).Name & "_"
            
            If Left(Name, Len(s)) = s Then
                IsEventHandler = True
                Exit Property
            End If
            
        End If
        
    Next i
    
End Property


Private Function GetSubName(SubHead As String) As String
    
    Dim StartPos As Long
    Dim EndPos As Long
    
    EndPos = InStr(1, SubHead, "(")
    StartPos = InStrRev(SubHead, " ", EndPos) + 1
    
    GetSubName = Mid(SubHead, StartPos, EndPos - StartPos)
    
End Function

Function GetSubScope(SubHead As String) As VBScopes
    
    Const SubScopesString As String = "Private|Public"
    
    Dim SubScopes() As String
    Dim i As Long
    
    SubScopes = Split(SubScopesString, "|")
    
    For i = 0 To UBound(SubScopes)
        If InText(1, SubHead, SubScopes(i)) <> 0 Then
            Select Case SubScopes(i)
                Case "Private"
                    GetSubScope = VBPrivate
                Case "Public"
                    GetSubScope = VBPublic
            End Select
            Exit Function
        End If
    Next i
    
    GetSubScope = VBPublic
    
End Function

Function GetSubType(SubHead As String) As String
    
    Const SubTypesString As String = "Sub|Function|Property Get|Property Let|Property Set"
    
    Dim SubTypes() As String
    Dim i As Long
    
    SubTypes = Split(SubTypesString, "|")
    
    For i = 0 To UBound(SubTypes)
        If InText(1, SubHead, SubTypes(i)) <> 0 Then
            GetSubType = SubTypes(i)
            Exit Function
        End If
    Next i
    
    
End Function

Private Function GetSubDef(SubCode) As String
    
    Dim Lines() As String
    Dim Line As String
    Dim LastLine As Long
    Dim s As String
    Dim tmp As String
    Dim i As Long
    
    'Debug.Print SubCode
    
    Lines = Split(SubCode, vbNewLine)
    If Right(RTrim(Lines(0)), 1) = "_" Then
        
        For i = 0 To UBound(Lines)
            Line = RTrim(Lines(i))
            If Right(Line, 1) <> "_" Then
                LastLine = i
                Exit For
            End If
        Next i
        
        For i = 0 To LastLine
            Line = RTrim(Lines(i))
            If i <> LastLine Then
                Line = Left(Line, Len(Line) - 1)
                Line = RTrim(LTrim(Line))
                If Right(s, 1) = "," Then
                    s = s & " "
                End If
                s = s & Line
            Else
                If Right(s, 1) = ")" Then
                    s = s & " "
                End If
                s = s & RTrim(LTrim(Line))
                Exit For
            End If
        Next i
    Else
        s = Lines(0)
    End If
    
    GetSubDef = s
    
End Function

Private Function FindParameters()
    
    Dim StartPos As Long
    Dim EndPos As Long
    Dim a() As String
    Dim s As String
    Dim i As Long
    
    s = HeadLine
    
    StartPos = InStr(1, s, "(") + 1
    EndPos = InStrRev(s, ")")
        
    s = Mid(s, StartPos, EndPos - StartPos)
    
    If s = "" Then Exit Function
    
    a() = Split(s, ",")
    'ReDim Params(1 To UBound(a) + 1)
    
    
    
    For i = 0 To UBound(a())
        
        ParameterCount = ParameterCount + 1
        Set Parameters(ParameterCount) = New VBParameter
        
        With Parameters(ParameterCount)
            Set .ParentSub = Me
            .ID = ParameterCount
            .Alias.AliasType = VBParameter + VBVariable
            .Alias.Scope = VBLocal
            .LoadDefination a(i)
        End With
        
    Next i
    
End Function

Function PreProcess()
    
   
    
End Function

Function CryptSub()
    
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim l As Long
    Dim r As Long
    Dim s As String
    
    Dim StartPos As Long
    Dim OldName As String
    Dim NewName As String
    
    If Not Me.IsEventHandler Then
        For i = 1 To ParameterCount
            With Parameters(i)
                ParentModule.BlockReplace Me.Block, .Alias.Name, .Alias.NewName
            End With
        Next i
    End If
    
    For i = 1 To VariableCount
        With Variables(i)
            ParentModule.BlockReplace Me.Block, .Alias.Name, .Alias.NewName
        End With
    Next i
    
    'Debug.Assert Name <> "MimeEncode"
    
    If CBool(Me.Alias.Scope And VBPublic) Then
        
        ParentModule.ParentProject.GlobalReplace Alias.Name, Alias.NewName
    
    Else
        
        ParentModule.ReplaceInSubs Alias.Name, Alias.NewName
    
    End If
    
End Function


'Function FindVariables()
'
'    Const LocalVarIndicators  As String = "Dim|Static|Const"
'
'    Dim i As Long
'    Dim l As Long
'    Dim r As Long
'    Dim c As Long
'    Dim n As Long
'
'    Dim ParentModule.Lines() As String
'    Dim Indicators() As String
'    Dim Parameters() As String
'
'    Dim VarLine As String
'    Dim VarsStr As String
'    Dim VarNames() As String
'
'    Indicators = Split(LocalVarIndicators, "|")
'    ParentModule.Lines = Split(Code, vbNewLine)
'
'    l = 0
'    Do
'
'        i = 0
'        Do
'
'            r = InText(1, ParentModule.Lines(l), Indicators(i))
'            If r <> 0 Then
'
'                VarsStr = Mid(ParentModule.Lines(l), r + Len(Indicators(i)) + 1)
'                VarNames = Split(VarsStr, ", ")
'
'                For n = 0 To UBound(VarNames)
'
'                    VarLine = Indicators(i) & " " & VarNames(n)
'
'                    VariableCount = VariableCount + 1
'                    Set Variables(VariableCount) = New VBVariable
'
'                    With Variables(VariableCount)
'
'                        Set .ParentSub = Me
'                        .ID = VariableCount
'                        .Alias.Scope = VBLocal
'                        .LoadLine VarLine
'
'                    End With
'
'                Next n
'
'            End If
'
'            i = i + 1
'        Loop While i <= UBound(Indicators)
'
'        l = l + 1
'    Loop While l <= UBound(ParentModule.Lines)
'
'End Function


'Private Function FindVariables()
'
'    Const DeclIndicatorsString      As String = "Dim|Static"
'
'    Const LongDataTypeIndicator     As String = "As"
'
'    Const TypeSignsString           As String = "@|#|%|&|!|$"
'    Const DataTypesString           As String = "Boolean|Byte|Currency|Date|Double|Integer|Long|Object|Single|String|Variant"
'
'    Const ConstIndicator            As String = "Const"
'
'
'    Dim DeclIndicators()            As String
'
'    Dim DataTypes()                 As String
'    Dim TypeSigns()                 As String
'
'    Dim VarsStr                     As String
'    Dim Vars()                      As String
'    Dim VarLine                     As String
'
'
'    Dim l As Long
'    Dim i As Long
'    Dim n As Long
'    Dim y As Long
'
'    Dim r As Long
'    Dim x As Long
'    Dim z As Long
'    Dim c As Long
'
'    Dim Found As Boolean
'    Dim IsConstant As Boolean
'
'
'    'If Trim(Code) = "" Then
'    '    Exit Function
'    'End If
'
'
'
'    DeclIndicators() = Split(DeclIndicatorsString, "|")
'
'    DataTypes() = Split(DataTypesString, "|")
'    TypeSigns() = Split(TypeSignsString, "|")
'
'
'    For l = Block.FirstLine + 1 To Block.LastLine - 1
'
'
'        If Trim(ParentModule.Lines(l)) <> "" Then
'
'            z = AnyInText(ParentModule.Lines(l), DeclIndicators())
'            If z <> 0 Then
'
'                r = InText(1, ParentModule.Lines(l), DeclIndicators(z - 1))
'
'
'
'                    VarsStr = Mid(ParentModule.Lines(l), r + Len(DeclIndicators(z - 1)) + 1)
'                    VarsStr = LTrim(VarsStr)
'                    x = InText(1, VarsStr, ConstIndicator)
'                    If x <> 0 Then
'                        IsConstant = True
'                        VarsStr = Mid(VarsStr, Len(ConstIndicator) + 1)
'                        VarsStr = LTrim(VarsStr)
'                    Else
'                        IsConstant = False
'                    End If
'                    Vars() = Split(VarsStr, ", ")
'
'
'                    For n = 0 To UBound(Vars)
'
'                        VarLine = DeclIndicators(z - 1) & " "
'                        If IsConstant Then
'                            VarLine = VarLine & "Const "
'                        End If
'                        VarLine = VarLine & Vars(n)
'
'                        VariableCount = VariableCount + 1
'                        Set Variables(VariableCount) = New VBVariable
'
'                        With Variables(VariableCount)
'
'                            Set .ParentSub = Me
'                            .ID = VariableCount
'                            .LoadLine VarLine
'
'                        End With
'
'
'                    Next n
'
'
'            End If
'
'        End If
'
'    Next l
'
'
'End Function

Private Function FindVariables()
    
    Const DeclIndicatorsString      As String = "Dim |Static "
    
    Const LongDataTypeIndicator     As String = "As"
    
    Const TypeSignsString           As String = "@|#|%|&|!|$"
    Const DataTypesString           As String = "Boolean|Byte|Currency|Date|Double|Integer|Long|Object|Single|String|Variant"
    
    Const ConstIndicator            As String = "Const"
    
    
    Dim DeclIndicators()            As String
    Dim DeclIndicator               As String
    
    Dim DataTypes()                 As String
    Dim TypeSigns()                 As String
    
    Dim VarsStr                     As String
    Dim Vars()                      As String
    Dim VarLine                     As String

    
    Dim l As Long
    Dim i As Long
    Dim n As Long
    Dim y As Long
    
    Dim r As Long
    Dim x As Long
    Dim z As Long
    Dim c As Long
    
    Dim Found As Boolean
    Dim IsConstant As Boolean
    
    
    'If Trim(Code) = "" Then
    '    Exit Function
    'End If
    
  
    
    DeclIndicators() = Split(DeclIndicatorsString, "|")
    
    DataTypes() = Split(DataTypesString, "|")
    TypeSigns() = Split(TypeSignsString, "|")
    
    
    For l = Block.FirstLine + 1 To Block.LastLine - 1
     
        
        If Trim(ParentModule.Lines(l)) <> "" Then
            
            z = LineBeginIsIn(ParentModule.Lines(l), DeclIndicators())
            
            If z <> -1 Then
                
                DeclIndicator = RTrim(DeclIndicators(z))
                
                r = InText(1, ParentModule.Lines(l), DeclIndicator)
                
                    
                    
                    VarsStr = Mid(ParentModule.Lines(l), r + Len(DeclIndicator) + 1)
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
                    
                        VarLine = DeclIndicator & " "
                        If IsConstant Then
                            VarLine = VarLine & "Const "
                        End If
                        VarLine = VarLine & Vars(n)
                    
                        VariableCount = VariableCount + 1
                        Set Variables(VariableCount) = New VBVariable
                    
                        With Variables(VariableCount)
                        
                            Set .ParentSub = Me
                            .ID = VariableCount
                            .LoadLine VarLine
                            
                        End With
                        
                    
                    Next n
                    
            
            End If
            
        End If
            
    Next l
    
    
End Function



Function GetSubDataType(SubDef As String) As String
    
    Dim s As String
    
    s = Left(SubDef, InStr(1, SubDef, "("))
    s = s & Mid(SubDef, InStrRev(SubDef, ")"))
    
    GetSubDataType = GetDataType(s)
    
End Function


Property Get GlobalName() As String
    GlobalName = ParentModule.Name & "." & Name
End Property

Property Get NewGlobalName() As String
    NewGlobalName = ParentModule.Alias.NewName & "." & Alias.NewName
End Property


Function Replace(OldWord As String, NewWord As String)
    Code = ReplaceWords(Code, OldWord, NewWord)
End Function



Function ReplaceEx(Expression As String, DataType As String)
    
    Dim i As Long
    
    For i = 1 To VariableCount
        
        With Variables(i)
            
            If .Alias.Name = Expression Then
                If .Alias.DataType = DataType Then
                    Code = ReplaceWords(Code, .Alias.Name, Expression)
                End If
            End If
            
        End With
        
    Next i
    
End Function


Property Get HasErrorHandler() As Boolean
    
    Dim i As Long
    Dim StrLine As String
    
    For i = Block.FirstLine To Block.LastLine
        
        StrLine = Me.ParentModule.Lines(i)
        
        If Left(LTrim(StrLine), Len("On Error")) = "On Error" Then
            
            HasErrorHandler = True
            Exit Property
            
        End If
        
    Next i
    
End Property

'Function AddErrorHandler()
'On Error GoTo OnErr
'
'    Dim NewLine As String
'    Dim SubType2 As String
'    Dim StrErrHandler As String
'
'    Dim x As Long
'
'    x = InStr(1, SubType, " ")
'
'    If x <> 0 Then
'        SubType2 = Left(SubType, x - 1)
'    Else
'        SubType2 = SubType
'    End If
'
'
'    StrErrHandler = fOptions.txtErrorHandler.Text
'    StrErrHandler = VBA.Replace(StrErrHandler, "<Module>", Me.ParentModule.Name)
'    StrErrHandler = VBA.Replace(StrErrHandler, "<Sub>", Me.Name)
'
'
'    NewLine = Me.HeadLine & vbNewLine & "On Error GoTo OnErr"
'
'    ParentModule.Lines(Me.Block.FirstLine) = NewLine
'
'    NewLine = "Exit " & SubType2 & vbNewLine _
'            & "OnErr: " & vbNewLine _
'            & StrErrHandler & vbNewLine _
'            & Me.ParentModule.Lines(Me.Block.LastLine)
'
'
'
'    ParentModule.Lines(Me.Block.LastLine) = NewLine
'
'Exit Function
'OnErr:
'    Debug.Print Err.Description
'    Resume Next
'End Function

Property Get EventName() As String
    
    Dim x As Long
    Dim s As String
    
    x = InStrRev(Name, "_")
    
    If x <> 0 Then
        
        EventName = Mid(Name, x + 1)
    
    End If
    
End Property
