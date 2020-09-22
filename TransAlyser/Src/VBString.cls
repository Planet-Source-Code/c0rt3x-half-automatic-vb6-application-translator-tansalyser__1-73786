VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ParentModule   As VBModule
Public ParentSub      As VBSub
Public GlobalID       As Long
Public LineID         As Long
Public Index          As Long
Public Value          As String

Private pTranslation    As String
Private pChecked        As Boolean


Private Sub Class_Initialize()
    'pChecked = True
End Sub

Public Property Get Checked() As Boolean
    
    Checked = pChecked

End Property

Public Property Let Checked(NewVal As Boolean)
    
    Dim i As Long
    Dim x As Long
    Dim y As Long
    
    pChecked = NewVal
        
    For i = 1 To My.Project.StringCount
        
        If My.Project.Strings(i).Value = Me.Value Then
            
            If My.Project.Strings(i).Checked <> NewVal Then
                My.Project.Strings(i).Checked = NewVal
            End If
            
        End If
        
    Next i
    
    For x = 1 To My.Project.ModuleCount
        
        For y = 1 To My.Project.Modules(x).StringCount
            
            If My.Project.Modules(x).Strings(y).Value = Me.Value Then
            
                If My.Project.Modules(x).Strings(y).Checked <> NewVal Then
                        
                    My.Project.Modules(x).Strings(y).Checked = NewVal
                        
                End If
                
            End If
            
        Next y
        
    Next x
    
        
End Property

Property Get Translation() As String
    Translation = pTranslation
End Property

Public Property Let Translation(NewVal As String)
    
    Dim i As Long
    Dim x As Long
    Dim y As Long
    
    pTranslation = NewVal
        
    For i = 1 To My.Project.StringCount
        
        If My.Project.Strings(i).Value = Me.Value Then
            
            If My.Project.Strings(i).Translation <> NewVal Then
                My.Project.Strings(i).Translation = NewVal
            End If
            
        End If
        
    Next i
    
    For x = 1 To My.Project.ModuleCount
        
        For y = 1 To My.Project.Modules(x).StringCount
            
            If My.Project.Modules(x).Strings(y).Value = Me.Value Then
            
                If My.Project.Modules(x).Strings(y).Translation <> NewVal Then
                        
                    My.Project.Modules(x).Strings(y).Translation = NewVal
                        
                End If
                
            End If
            
        Next y
        
    Next x
    
        
End Property



Function EncryptString()
    
    Dim DecStr As String
    Dim EncStr As String
    
    Dim FuncName As String
    
    'FuncName = "VBStringBase64Decode"
    'Debug.Print ParentModule.ParentProject.Modules(1).Subs(1).Alias.Name
    
'    If My.Config.EncryptStrings Then
'
'        FuncName = ParentModule.ParentProject.Modules(ParentModule.ParentProject.SpecialModuleID(StringDecryptionModule)).Subs(1).Alias.NewName
'
'        DecStr = Chr(34) & Value & Chr(34)
'
'        EncStr = FuncName & "(" & Chr(34) & ParentModule.ParentProject.StringEncryption.Base64Encode(Value) & Chr(34) & ")"
'
'        ParentModule.Lines(LineID) = Replace(ParentModule.Lines(LineID), DecStr, EncStr)
'
'    End If
'
'    If (My.Config.ExportStrings Or My.Config.EnableCrackCheck) Then
'
'        If My.Config.ExportStrings Then
'            FuncName = ParentModule.ParentProject.Modules(ParentModule.ParentProject.SpecialModuleID(StringLoadingModule)).Subs(1).Alias.NewName
'        Else
'            FuncName = ParentModule.ParentProject.Modules(ParentModule.ParentProject.SpecialModuleID(CrackCheckModule)).Subs(1).Alias.NewName
'        End If
'
'        DecStr = Chr(34) & Value & Chr(34)
'
'        EncStr = FuncName & "(" & CStr(Me.GlobalID) & ")"
'
'        ParentModule.Lines(LineID) = Replace(ParentModule.Lines(LineID), DecStr, EncStr)
'
'    End If
'
End Function

Function ReplaceString()
    
    
    Dim OldBlock As String
    Dim NewBlock As String
    
    Dim DecStr As String
    Dim EncStr As String
    
    Dim FuncName As String
    
    Dim x As Long
    Dim sLine As String
    
    
    Const DebugKeyword As String = "Benutzername"
    
    
    If Me.ParentSub Is Nothing Then
        
        Debug.Assert False
        
    End If
    
    FuncName = "GetTxt"
    
    DecStr = Chr(34) & Value & Chr(34)
    EncStr = FuncName & "(" & Chr(34) & Value & Chr(34) & ")"
    
    OldBlock = Mid(Me.ParentModule.FileData, Me.ParentModule.SubStartPos)
    NewBlock = Mid(Me.ParentModule.FileData, Me.ParentModule.SubStartPos)
    
   
    
    
    If InStr(1, Me.Value, "Benutzer", vbTextCompare) > 0 Then
    
         
        x = InStr(1, OldBlock, Me.Value, vbTextCompare)
        
        If x > 0 Then
        
            sLine = GetLine(OldBlock, x)
            
           
         
        End If
    
    
        If InStr(1, Me.Value, DebugKeyword, vbTextCompare) > 0 Then
            
             Debug.Print sLine
            
            Debug.Assert False
            
            
        End If
        
        'NewBlock = ReplaceWords(NewBlock, DecStr, EncStr)
        
        'Debug.Print GetLine(OldBlock, InStr(1, OldBlock, "Benutzername", vbTextCompare))
        
    End If
    
    
    
    NewBlock = ReplaceWords(NewBlock, DecStr, EncStr) ', "GetTxt(")
    
    'NewBlock = Replace(NewBlock, DecStr, EncStr)
    
    If InStr(1, Me.Value, DebugKeyword, vbTextCompare) > 0 Then
    
        Debug.Print GetLine(NewBlock, InStr(1, NewBlock, "Benutzername", vbTextCompare))
        
        Debug.Assert False
    
    End If
    
    
    
    ParentModule.FileData = Replace(ParentModule.FileData, OldBlock, NewBlock)
    

    
End Function

Function LoadValue(NewValue As String)
    
    Dim i As Long
    
    'Value = Replace(NewValue, Chr(34) & Chr(34), Chr(34))
    Value = NewValue
    
    i = ParentModule.ParentProject.StringIndex(NewValue)
    
    If i = 0 Then
    
        ParentModule.ParentProject.StringCount = ParentModule.ParentProject.StringCount + 1
        GlobalID = ParentModule.ParentProject.StringCount
        
        Set ParentModule.ParentProject.Strings(GlobalID) = Me
    
    Else
        
        GlobalID = i
        
        Set ParentModule.ParentProject.Strings(GlobalID) = Me
        
    End If
    
End Function


