VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBEnumValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParentEnum           As VBEnum
Private pAlias                As VBAlias
Private pSrcLine              As String
Private pValue                As Long

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


Property Get ParentEnum() As VBEnum
    Set ParentEnum = pParentEnum
End Property

Property Set ParentEnum(NewObject As VBEnum)
    Set pParentEnum = NewObject
    Set pAlias.ParentProject = pParentEnum.ParentModule.ParentProject
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


Property Get Value() As String
    Value = pValue
End Property

Property Let Value(NewValue As String)
    pValue = NewValue
End Property


Property Get SrcLine() As String
    SrcLine = pSrcLine
End Property

Property Let SrcLine(NewValue As String)
    pSrcLine = NewValue
End Property


Function LoadLine(LineStr As String)

    Dim x As Long
    Dim s As String
    Dim v As String
    
    pSrcLine = LineStr
    
    s = LineStr
    s = LTrim(RTrim(s))
    
    x = InStr(1, s, "=")
    
    If x <> 0 Then
        Name = LTrim(RTrim(Mid(s, 1, InStr(1, s, "=") - 1)))
    Else
        Name = s
    End If
    
    
    If InStr(1, Name, "[") <> 0 Then
        Name = Mid(s, 2, Len(Name) - 2)
    End If
     
       
    If x <> 0 Then
        v = LTrim(RTrim((Mid(s, x + 1))))
        If IsNumber(v) Then
            pValue = CLng(v)
        Else
            pValue = ParentEnum.GetValue(v)
        End If
    Else
        If ParentEnum.ValueCount > 1 Then
            Value = ParentEnum.Values(ParentEnum.ValueCount - 1).Value + 1
        Else
            pValue = 0
        End If
    End If
    
End Function











