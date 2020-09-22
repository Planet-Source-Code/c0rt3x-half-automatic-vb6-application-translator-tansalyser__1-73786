VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBAlias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ParentProject          As VBProject

Public ModuleID               As Long
Public ControlID              As Long
Public DeclarationID          As Long
Public EventID                As Long
Public TypeID                 As Long
Public EnumID                 As Long
Public SubID                  As Long
Public ParameterID            As Long
Public VariableID             As Long

Public ID                     As Long
Private pName                 As String
Public NewName                As String
Public DataType               As String
Public AliasType              As VBAliasTypes
Public Scope                  As VBScopes
Public Priority               As Long
Public Protected              As Boolean

Public Property Get ParentModule() As VBModule
    Set ParentModule = ParentProject.Modules(ModuleID)
End Property

Public Property Get ParentControl() As VBControl
    Set ParentControl = ParentModule.Controls(ControlID)
End Property

Public Property Get ParentSub() As VBSub
    Set ParentSub = ParentModule.Subs(SubID)
End Property

Property Get ParentVariable() As VBVariable

    If (Scope And VBPrivate) Or (Scope And VBPublic) Then
        Set ParentVariable = ParentModule.Variables(VariableID)
    ElseIf (Scope And VBLocal) Then
        ParentVariable = ParentModule.Subs(SubID).Variables(VariableID)
    End If

End Property

Public Property Get ParentParameter() As VBParameter
    Set ParentParameter = ParentSub.Parameters(ParameterID)
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(NewVal As String)
    If NewVal = "" Then
        Debug.Assert False
    End If
    pName = NewVal
End Property
