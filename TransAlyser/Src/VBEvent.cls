VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParentModule         As VBModule
Private pAlias                As VBAlias

Private pLineID               As Long
Private pCode                 As String

Private pParameterCount       As Long
Private pParameters()         As VBParameter

Private Sub Class_Initialize()
    
    Set pAlias = New VBAlias
    
    With pAlias
        .AliasType = VBDeclaration
    End With
    
    ReDim pParameters(0)
    
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
    ID = pAlias.EventID
End Property

Property Let ID(NewValue As Long)
    pAlias.EventID = NewValue
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


Function LoadLine(LineID As Long)
    
    
    Dim LineStr As String
    Dim i As Long
    Dim x As Long
    Dim KeyWord As String
    Dim a As Long, z As Long
    Dim s As String
    
    pLineID = LineID
    LineStr = ParentModule.Lines(LineID)
    Code = LineStr
    
    KeyWord = GetWord(1, LineStr)
    
    Me.Alias.Scope = SelectScope(KeyWord)
    
    
    x = InText(1, LineStr, "Event")
    x = x + Len("Event") + 1
    
    Name = GetWord(x, LineStr)
    
    FindParameters
    
End Function




Private Function FindParameters()
    
    Dim StartPos As Long
    Dim EndPos As Long
    Dim a() As String
    Dim s As String
    Dim i As Long
    
    s = pCode
    
    'Debug.Print s
    
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
            Set .ParentEvent = Me
            .ID = ParameterCount
            .Alias.AliasType = VBParameter + VBVariable
            .Alias.Scope = VBLocal
            .LoadDefination a(i)
        End With
        
    Next i
    
End Function


Function Rebuild() As String
    
    Dim s As String
    Dim i As Long
    
    
    If (Alias.Scope And VBPublic) Then
        s = "Public "
    ElseIf (Alias.Scope And VBPrivate) Then
        s = "Private "
    End If
    
    s = s & "Event " & Alias.NewName & " " & "("
    
    
    For i = 1 To ParameterCount
        
        s = s & Parameters(i).Rebuild
        
        If i <> ParameterCount Then
            s = s & ", "
        End If
        
    Next i
    
    s = s & ")"
    
    Rebuild = s
    
End Function



Function Crypt()

    ParentModule.Lines(pLineID) = Me.Rebuild
    
    If (Me.Alias.Scope And VBPublic) Then
        ParentModule.ParentProject.GlobalReplace Alias.Name, Alias.NewName
    Else
        ParentModule.ReplaceInSubs Alias.Name, Alias.NewName
    End If

End Function




