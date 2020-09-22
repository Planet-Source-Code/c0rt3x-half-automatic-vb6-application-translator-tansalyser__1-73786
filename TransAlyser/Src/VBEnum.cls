VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParentModule           As VBModule
Private pAlias                  As VBAlias
Private pCode                   As String

Private pBlock                  As VBCodeBlock

Private pValueCount            As Long
Private pValues()              As VBEnumValue

Private Sub Class_Initialize()
    
    Set pAlias = New VBAlias
    Set pBlock = New VBCodeBlock
    
    With pAlias
        .AliasType = VBType
    End With
    
    ReDim pValues(0)
    
End Sub


Property Get ParentModule() As VBModule
    Set ParentModule = pParentModule
End Property

Property Set ParentModule(NewObject As VBModule)
    Set pParentModule = NewObject
    Set pAlias.ParentProject = pParentModule.ParentProject
    pAlias.ModuleID = NewObject.ID
    pAlias.EnumID = ID
End Property

Property Get Alias() As VBAlias
    Set Alias = pAlias
End Property


Property Get ID() As Long
    ID = pAlias.EnumID
End Property

Property Let ID(NewValue As Long)
    pAlias.EnumID = NewValue
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


Property Get ValueCount() As Long
    ValueCount = pValueCount
End Property

Property Let ValueCount(NewValue As Long)
    ReDim Preserve pValues(NewValue)
    pValueCount = NewValue
End Property

Property Get Values(Index As Long) As VBEnumValue
    Set Values = pValues(Index)
End Property

Property Set Values(Index As Long, NewObject As VBEnumValue)
    Set pValues(Index) = NewObject
End Property


Property Get GetValue(Name As String) As Long
    Dim i As Long
    For i = 1 To ValueCount
        If Values(i).Name = Name Then
            GetValue = Values(i).Value
            Exit Property
        End If
    Next i
End Property

Sub LoadEnumCode(FirstLine As Long, LastLine As Long)
     
    Dim Lines() As String
    Dim i As Long
    Dim x As Long
    Dim s As String
    Dim a As Long, z As Long
    
    'pCode = EnumCode
    'pBlock.Code = EnumCode
    
    pBlock.FirstLine = FirstLine
    pBlock.LastLine = LastLine
    
    
    s = LTrim(RTrim(ParentModule.Lines(FirstLine)))
    s = GetWord(1, s)
    If s <> "Enum" Then
        Me.Alias.Scope = SelectScope(s)
    Else
        Me.Alias.Scope = VBLocal + VBPrivate + VBPublic
    End If
    
    
    s = LTrim(RTrim(ParentModule.Lines(FirstLine)))
    x = InText(1, s, "Enum")
    a = x + Len("Enum") + 1
    s = Mid(s, a)
    
    If s = "" Then
        Debug.Assert False
    End If
    
    Name = s
    
    
    For i = FirstLine + 1 To LastLine - 1
        
        If LTrim(RTrim(ParentModule.Lines(i))) <> "" Then
        
            ValueCount = ValueCount + 1
            Set Values(ValueCount) = New VBEnumValue
            With Values(ValueCount)
                Set .ParentEnum = Me
                .LoadLine ParentModule.Lines(i)
            End With
        
        End If
        
    Next i
    

    
End Sub


Function Crypt()
    
    Dim i As Long
   
    ParentModule.ReplaceInMod Alias.Name, Alias.NewName
    
    If Alias.Scope And VBPublic Then
    
        ParentModule.ParentProject.GlobalReplace Alias.Name, Alias.NewName
        
        For i = 1 To ValueCount
            ParentModule.ParentProject.GlobalReplace Hug(Values(i).Alias.Name), Hug(Values(i).Alias.NewName)
            ParentModule.ParentProject.GlobalReplace Values(i).Alias.Name, Values(i).Alias.NewName
        Next i
        
    Else

        For i = 1 To ValueCount
            ParentModule.ReplaceInMod Hug(Values(i).Alias.Name), Hug(Values(i).Alias.NewName)
            ParentModule.ReplaceInMod Values(i).Alias.Name, Values(i).Alias.NewName
        Next i

    End If
    
    
End Function
