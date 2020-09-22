VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParentModule           As VBModule
Private pAlias                  As VBAlias
Private pCode                   As String

Private pBlock                  As VBCodeBlock

Private pMemberCount            As Long
Private pMembers()              As VBTypeMember

Private Sub Class_Initialize()
    
    Set pAlias = New VBAlias
    Set pBlock = New VBCodeBlock
    
    
    With pAlias
        .AliasType = VBType
    End With
    
    ReDim pMembers(0)
    
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
    ID = pAlias.TypeID
End Property

Property Let ID(NewValue As Long)
    pAlias.TypeID = NewValue
End Property

Property Get Name() As String
    Name = pAlias.Name
End Property

Property Let Name(NewValue As String)
    pAlias.Name = NewValue
End Property

Property Get Code() As String
    Code = pCode
End Property

Property Let Code(NewValue As String)
    pCode = NewValue
End Property


Property Get Block() As VBCodeBlock
    Set Block = pBlock
End Property

Property Let Block(NewObject As VBCodeBlock)
    Set pBlock = NewObject
End Property


Property Get MemberCount() As Long
    MemberCount = pMemberCount
End Property

Property Let MemberCount(NewValue As Long)
    ReDim Preserve pMembers(NewValue)
    pMemberCount = NewValue
End Property

Property Get Members(Index As Long) As VBTypeMember
    Set Members = pMembers(Index)
End Property

Property Set Members(Index As Long, NewObject As VBTypeMember)
    Set pMembers(Index) = NewObject
End Property


Sub LoadTypeCode(FirstLine As Long, LastLine As Long)
     
    Dim Lines() As String
    Dim i As Long
    Dim x As Long
    Dim s As String
    Dim a As Long, z As Long
    
    'Code = TypeCode
    'Block.Code = TypeCode
    
    Block.FirstLine = FirstLine
    Block.LastLine = LastLine
    
    
    'Lines = Split(TypeCode, vbNewLine)
    
    s = LTrim(RTrim(ParentModule.Lines(FirstLine)))
    s = GetWord(1, s)
    If s <> "Type" Then
        Me.Alias.Scope = SelectScope(s)
    Else
        Me.Alias.Scope = VBLocal + VBPrivate + VBPublic
    End If
    
    s = LTrim(RTrim(ParentModule.Lines(FirstLine)))
    x = InText(1, s, "Type")
    a = x + Len("Type") + 1
    s = Mid(s, a)
    Name = s
    
    
    For i = FirstLine + 1 To LastLine - 1
        
        If Trim(ParentModule.Lines(i)) <> "" Then
        
            MemberCount = MemberCount + 1
            Set Members(MemberCount) = New VBTypeMember
            With Members(MemberCount)
                Set .ParentType = Me
                .ID = MemberCount
                .LoadLine ParentModule.Lines(i)
            End With
        
        End If
        
    Next i
    
    
End Sub

Function Crypt()
    
    Dim i As Long
    
'    If Me.Name = "VBQuote" Then
'        Debug.Assert False
'    End If
    
    ParentModule.ReplaceInMod Alias.Name, Alias.NewName
    
    If Alias.Scope And VBPublic Then
    
        ParentModule.ParentProject.GlobalReplace Alias.Name, Alias.NewName
        
        For i = 1 To MemberCount
            ParentModule.ParentProject.GlobalReplace Members(i).Alias.Name, Members(i).Alias.NewName
        Next i
        
    Else

        For i = 1 To MemberCount
            ParentModule.ReplaceInMod Members(i).Alias.Name, Members(i).Alias.NewName
        Next i

    End If
    
    
End Function
