VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParentDeclaration    As VBDeclaration
Private pParentEvent          As VBEvent
Private pParentSub            As VBSub
Private pAlias                As VBAlias

Private pByReference          As Boolean
Private pByValue              As Boolean
Private pIsOptional           As Boolean
Private pIsArray              As Boolean
Private pDefaultValue         As String

Private Const DataNameList As String = "Byte|Boolean|Integer|Long|Single|Double|Currency|Date|String|Object|Variant"
Private Const DataSymbolList As String = "||%|&|!|#|@||$"

Private Sub Class_Initialize()
    
    Set pAlias = New VBAlias
    
    With pAlias
        .AliasType = VBParameter
        .Scope = VBLocal
    End With
    
End Sub


Property Get ParentSub() As VBSub
    Set ParentSub = pParentSub
End Property

Property Set ParentSub(NewObject As VBSub)
    Set pParentSub = NewObject
    Set pAlias.ParentProject = pParentSub.ParentModule.ParentProject
    pAlias.SubID = NewObject.ID
    pAlias.ModuleID = NewObject.ParentModule.ID
End Property


Property Get ParentDeclaration() As VBDeclaration
    Set ParentDeclaration = pParentDeclaration
End Property

Property Set ParentDeclaration(NewObject As VBDeclaration)
    Set pParentDeclaration = NewObject
    Set pAlias.ParentProject = NewObject.ParentModule.ParentProject
    pAlias.DeclarationID = NewObject.ID
    pAlias.ModuleID = NewObject.ParentModule.ID
End Property

Property Get ParentEvent() As VBDeclaration
    Set ParentEvent = pParentEvent
End Property

Property Set ParentEvent(NewObject As VBEvent)
    Set pParentEvent = NewObject
    Set pAlias.ParentProject = NewObject.ParentModule.ParentProject
    pAlias.EventID = NewObject.ID
    pAlias.ModuleID = NewObject.ParentModule.ID
End Property



Property Get Alias() As VBAlias
    Set Alias = pAlias
End Property


Property Get ID() As Long
    ID = pAlias.ParameterID
End Property

Property Let ID(NewValue As Long)
    pAlias.ParameterID = NewValue
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


Property Get ByReference() As Boolean
    ByReference = pByReference
End Property

Property Let ByReference(NewValue As Boolean)
    pByReference = NewValue
End Property


Property Get ByValue() As Boolean
    ByValue = pByValue
End Property

Property Let ByValue(NewValue As Boolean)
    pByValue = NewValue
End Property


Property Get IsOptional() As Boolean
    IsOptional = pIsOptional
End Property

Property Let IsOptional(NewValue As Boolean)
    pIsOptional = NewValue
End Property


Property Get IsArray() As Boolean
    IsArray = pIsArray
End Property

Property Let IsArray(NewValue As Boolean)
    pIsArray = NewValue
End Property


Property Get DefaultValue() As String
    DefaultValue = pDefaultValue
End Property

Property Let DefaultValue(NewValue As String)
    pDefaultValue = NewValue
End Property


Public Function LoadDefination(Def As String)
    
    Dim x As Long
    Dim s As String
    
    If InText(1, Def, "Optional") <> 0 Then
        
        IsOptional = True
        Def = ReplaceWords(Def, "Optional", "")
        
        x = InStr(1, Def, "=", vbBinaryCompare)
        If x <> 0 Then
            DefaultValue = RTrim(LTrim(Mid(Def, x)))
        End If
        
    End If
        
    If InText(1, Def, "ByRef") <> 0 Then
        ByReference = True
        Def = ReplaceWords(Def, "ByRef", "")
    End If
    
    If InText(1, Def, "ByVal") <> 0 Then
        ByValue = True
        Def = ReplaceWords(Def, "ByVal", "")
    End If
    
    Def = LTrim(Def)
    
    Name = GetWord(1, Def)
    
    If InStr(1, Def, "(") <> 0 Then
        IsArray = True
    End If
    
    DataType = GetDataType(Def)
    
End Function


Public Function Rebuild() As String
    
    Dim s As String
    
    If ByReference Then
        s = "ByRef "
    End If
    If ByValue Then
        s = "ByVal "
    End If
    
    s = s & Alias.NewName
    
    'Debug.Print Alias.Name
    Debug.Assert Alias.NewName <> ""
    
    If IsArray Then
        s = s & "()"
    End If
    
    s = s & " As " & DataType
    
    Rebuild = s
    
End Function
















