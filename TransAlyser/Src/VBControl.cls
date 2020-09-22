VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pParentModule     As VBModule
Private pAlias            As VBAlias
Private pBlock            As VBCodeBlock
Private pCode             As String
Private pLines()          As String

Private pPropertyCount     As Long
Private pProperties()      As VBProperty

Private Sub Class_Initialize()
    
    Set pAlias = New VBAlias
    Set pBlock = New VBCodeBlock
    
    With pAlias
        .AliasType = VBControl
        .Scope = VBLocal + VBPrivate
    End With
    
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

Property Get Block() As VBCodeBlock
    Set Block = pBlock
End Property

Property Get ID() As Long
    ID = pAlias.ControlID
End Property

Property Let ID(NewValue As Long)
    pAlias.ControlID = NewValue
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


Property Get GetPropertyIndex(PropertyName As String) As Long
    
    Dim i  As Long
    
    For i = 1 To PropertyCount
        
        If Properties(i).Name = PropertyName Then
            
            GetPropertyIndex = i
            
            Exit Property
            
        End If
        
    Next i
    
End Property

Function LoadControlCode(FirstLine As Long, LastLine As Long)
    
    Dim c As Long
    Dim i As Long
    Dim Words() As String
    Dim LineStr As String
    Dim PropStr As String
    
    'Debug.Print CtlCode
    
    pBlock.FirstLine = FirstLine
    pBlock.LastLine = LastLine
    pBlock.LineCount = LastLine - FirstLine
    
    'pCode = CtlCode
    'Block.Code = CtlCode
    
    LineStr = RTrim(LTrim(ParentModule.Lines(FirstLine)))
    Words = Split(LineStr, " ")
                
    Name = Words(2)
    DataType = Words(1)
    
    If Left(DataType, 3) = "VB." Then
        pBlock.Protected = True
    End If
    
    If pBlock.LineCount < 3 Then
        Exit Function
    End If
    
    c = FirstLine + 1
    Do
        
        If pBlock.Protected Then
            ParentModule.LineProtected(c) = True
        End If
        
        If InStr(1, ParentModule.Lines(c), " Begin ", vbTextCompare) <> 0 Then
            Exit Do
        End If
        
        PropertyCount = PropertyCount + 1
        Set Properties(PropertyCount) = New VBProperty
        Set Properties(PropertyCount).ParentModule = Me.ParentModule
        Set Properties(PropertyCount).ParentControl = Me
        
        With Properties(PropertyCount)
            
            If InText(1, ParentModule.Lines(c), "BeginProperty") = 0 Then
                If InText(1, ParentModule.Lines(c), "EndProperty") = 0 Then
                    .LoadLine c
                End If
            End If
            
            'For i = FirstLine To LastLine
            '    Debug.Print ParentModule.Lines(i)
            'Next i
            
            If InText(1, ParentModule.Lines(c), "BeginProperty") <> 0 Then
            
                For i = c + 1 To LastLine - 1
                
                    If InText(1, ParentModule.Lines(i), "EndProperty") = 0 Then
                        
                        .ChildCount = .ChildCount + 1
                        Set .Childs(.ChildCount) = New VBProperty
                        Set .Childs(.ChildCount).ParentModule = Me.ParentModule
                        .Childs(.ChildCount).LoadLine i
                        
                    Else
                        
                        c = i
                        Exit For
                
                    End If
                    
                    
                Next i
            
            End If
        
        End With
        
        c = c + 1
    Loop While c < LastLine
    
End Function


Function CryptControl()
    
    Dim LineStr As String
    Dim Owner As String
    Dim GlobalName As String
    Dim NewTypeName As String
    
    Dim i As Long
    Dim y As Long
    Dim n As Long
    
    
    If Not Alias.Protected Then
        
        Owner = ParentModule.ParentProject.Name & "."
        
        If Left(DataType, Len(Owner)) = Owner Then
            For i = 1 To ParentModule.ParentProject.ModuleCount
                If ParentModule.ParentProject.Modules(i).DataType = "UserControl" Then
                    
                    GlobalName = ParentModule.ParentProject.Name & "." & ParentModule.ParentProject.Modules(i).Name
                    
                    If GlobalName = Me.DataType Then
                        
                        
                        NewTypeName = ParentModule.ParentProject.Name & "." & ParentModule.ParentProject.Modules(i).Alias.NewName
                        
                        LineStr = ReplaceWords(ParentModule.Lines(pBlock.FirstLine), Me.DataType, NewTypeName)
                        ParentModule.Lines(pBlock.FirstLine) = LineStr
                        
                        With ParentModule.ParentProject.Modules(i)
                        
                            For y = 1 To PropertyCount
                                
                                n = .SubID(Properties(y).Name)
                                    
                                If n <> 0 Then
                                    Properties(y).Name = .Subs(n).Alias.NewName
                                    Properties(y).Crypt
                                End If
                                
                            Next y
                            
                        End With
                        
                    End If
                    
                End If
            Next i
            
        Else
            
            
            For y = 1 To PropertyCount
                Properties(y).Crypt
            Next y
            
        
        End If
        
        'ParentModule.BlockReplace ParentModule.Blocks(B_Properties), Alias.Name, Alias.NewName, True
        ParentModule.ParentProject.GlobalReplace Alias.Name, Alias.NewName
        'LineStr = ReplaceWords(ParentModule.Lines(pBlock.FirstLine), Alias.Name, Alias.NewName)
        'ParentModule.Lines(pBlock.FirstLine) = LineStr
        
        'LineStr = ReplaceWords(ParentModule.Lines(pBlock.FirstLine), Alias.Name, Alias.NewName)
        'ParentModule.Lines(pBlock.FirstLine) = LineStr
    
    End If
    
End Function

Public Property Get Index() As String
    
    Dim i As Long
    
    For i = 1 To pPropertyCount
        If pProperties(i).Name = "Index" Then
            Index = pProperties(i).Value
            Exit Property
        End If
    Next i
    
End Property
