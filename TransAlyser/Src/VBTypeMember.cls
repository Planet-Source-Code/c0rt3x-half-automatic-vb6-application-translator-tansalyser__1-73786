VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBTypeMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParentType           As VBType

Private pAlias                As VBAlias

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


Property Get ParentType() As VBSub
    Set ParentType = pParentType
End Property

Property Set ParentType(NewObject As VBType)
    Set pParentType = NewObject
    Set pAlias.ParentProject = pParentType.ParentModule.ParentProject
    pAlias.TypeID = NewObject.ID
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


Property Get DataType() As String
    DataType = pAlias.DataType
End Property

Property Let DataType(NewValue As String)
    pAlias.DataType = NewValue
End Property


Property Get VarLine() As String
    VarLine = pVarLine
End Property

Property Let VarLine(NewValue As String)
    pVarLine = NewValue
End Property


Function LoadLine(DimLine As String)

    Dim r As Long
    Dim VarStr As String
    Dim KeyWord As String
    
    pVarLine = DimLine
    VarStr = DimLine
    
    VarStr = LTrim(RTrim(VarStr))
    
    Name = GetWord(1, VarStr)
    
    DataType = GetDataType(DimLine)
    
End Function










