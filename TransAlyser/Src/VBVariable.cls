VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParentModule         As VBModule
Private pParentSub            As VBSub

Private pAlias                As VBAlias

Private pHasEvents            As Boolean
Private pIsConstant           As Boolean
Private pIsStatic             As Boolean
Private pDefaultValue         As String
Private pVarLine              As String


Private Sub Class_Initialize()
    
    Set pAlias = New VBAlias
    With pAlias
        .AliasType = VBVariable
        '.Scope = VBLocal + VBPrivate
    End With
    
End Sub


Property Get Alias() As VBAlias
    Set Alias = pAlias
End Property


Property Get ParentSub() As VBSub
    Set ParentSub = pParentSub
End Property

Property Set ParentSub(NewObject As VBSub)
    Set pParentSub = NewObject
    Set pAlias.ParentProject = pParentSub.ParentModule.ParentProject
    pAlias.SubID = NewObject.ID
End Property


Property Get ParentModule() As VBModule
    Set ParentModule = pParentModule
End Property

Property Set ParentModule(NewObject As VBModule)
    Set pParentModule = NewObject
    Set pAlias.ParentProject = pParentModule.ParentProject
    pAlias.ModuleID = NewObject.ID
End Property


Property Get ID() As Long
    ID = pAlias.VariableID
End Property

Property Let ID(NewValue As Long)
    pAlias.VariableID = NewValue
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


Property Get IsStatic() As Boolean
    IsStatic = pIsStatic
End Property

Property Let IsStatic(NewValue As Boolean)
    pIsStatic = NewValue
End Property


Property Get VarLine() As String
    VarLine = pVarLine
End Property

Property Let VarLine(NewValue As String)
    pVarLine = NewValue
End Property


Property Get IsConstant() As Boolean
    IsConstant = pIsConstant
End Property

Property Let IsConstant(NewValue As Boolean)
    pIsConstant = NewValue
End Property


Property Get HasEvents() As Boolean
    HasEvents = pHasEvents
End Property


Property Get DefaultValue() As String
    DefaultValue = pDefaultValue
End Property

Property Let DefaultValue(NewValue As String)
    pDefaultValue = NewValue
End Property


Function LoadLine(DimLine As String)

    Dim r As Long
    Dim VarStr As String
    Dim KeyWord As String
    
    pVarLine = DimLine
    VarStr = DimLine
    
    KeyWord = GetWord(1, DimLine)
    
    VarStr = LTrim(ReplaceWords(VarStr, KeyWord, ""))
    
    r = InText(1, DimLine, "Const")
    If r <> 0 Then
        pIsConstant = True
        DefaultValue = GetValue(VarStr)
        VarStr = LTrim(ReplaceWords(VarStr, "Const", ""))
    End If
    
    
    r = InText(1, DimLine, "WithEvents")
    If r <> 0 Then
        pHasEvents = True
        VarStr = LTrim(ReplaceWords(VarStr, "WithEvents", ""))
    End If
    
    
    Name = GetWord(1, VarStr)
    
    DataType = GetDataType(DimLine)
    Alias.Scope = SelectScope(KeyWord)
    
End Function


Private Function GetValue(VarStr As String) As String
    
    Dim s As String
    Dim x As Long
    Dim StartPos As Long
    Dim EndPos As Long
    
    s = RTrim(LTrim(VarStr))
    
    x = InStr(1, s, Chr(34))
    If x <> 0 Then
        StartPos = x + 1
        EndPos = InStrRev(s, Chr(34))
    Else
        StartPos = InStrRev(s, " = ") + Len(" = ")
        EndPos = Len(s) + 1
    End If
    
    GetValue = Mid(VarStr, StartPos, (EndPos - StartPos))
    
End Function











