VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private IniData As String

Public Translate                As Boolean

Public Obfuscate                As Boolean

Public ReplaceFileNames         As Boolean
Public ReplaceProjectNames      As Boolean
Public ReplaceModuleNames       As Boolean

Public RemoveCodeLayout         As Boolean

Public RemoveComments           As Boolean
Public RemoveUnderscores        As Boolean
Public RemoveColons             As Boolean
Public RemoveEmptyLines         As Boolean
Public RemoveSpaces             As Boolean

Public EncryptControlCaptions   As Boolean
Public EncryptStrings           As Boolean
Public ExportStrings            As Boolean
Public EnableCrackCheck         As Boolean

Public EnableDebugMode          As Boolean
Public AddErrorHandlers         As Boolean

Private Sub Class_Initialize()
    
    Translate = True
    
'    Me.Obfuscate = CBool(fOptions.chkObfuscate.Value)
'
'        Me.ReplaceFileNames = CBool(fOptions.chkReplaceFileNames.Value)
'        Me.ReplaceProjectNames = CBool(fOptions.chkReplaceProjectNames.Value)
'        Me.ReplaceModuleNames = CBool(fOptions.chkReplaceModuleNames.Value)
'
'
'    Me.RemoveCodeLayout = CBool(fOptions.chkRemoveCodeLayout.Value)
'
'        Me.RemoveComments = CBool(fOptions.chkRemoveComments.Value)
'        Me.RemoveUnderscores = CBool(fOptions.chkRemoveUnderscores.Value)
'
'        Me.RemoveColons = CBool(fOptions.chkRemoveColons.Value)
'        Me.RemoveEmptyLines = CBool(fOptions.chkRemoveEmptyLines.Value)
'        Me.RemoveSpaces = CBool(fOptions.chkRemoveSpaces.Value)
'
'
'    Me.EncryptStrings = CBool(fOptions.chkEncryptStrings.Value)
'
'        Me.EncryptControlCaptions = CBool(fOptions.chkEncryptControlCaptions.Value)
'        Me.ExportStrings = CBool(fOptions.chkExportStrings.Value)
'        Me.EnableCrackCheck = CBool(fOptions.chkCrackCheck.Value)
'
'
'
'
'    Me.EnableDebugMode = CBool(fOptions.chkDebugMode.Value)
'
'        Me.AddErrorHandlers = CBool(fOptions.chkAddErrorHandlers.Value)
    
    
End Sub

Property Get IniPath()
    
    IniPath = My.Paths.AppPath & App.ExeName & ".ini"
    'IniPath = FixPath(App.Path) & App.ExeName & ".ini"
    
End Property

Function SaveConfig()
    
    SetIniValue IniData, "Config", "Obfuscate", Abs(CLng(Me.Obfuscate))
    
    SetIniValue IniData, "Config", "ReplaceFileNames", Abs(CLng(Me.ReplaceFileNames))
    SetIniValue IniData, "Config", "ReplaceProjectNames", Abs(CLng(Me.ReplaceProjectNames))
    SetIniValue IniData, "Config", "ReplaceModuleNames", Abs(CLng(Me.ReplaceModuleNames))
    
    SetIniValue IniData, "Config", "RemoveCodeLayout", Abs(CLng(Me.RemoveCodeLayout))
    
    SetIniValue IniData, "Config", "RemoveComments", Abs(CLng(Me.RemoveComments))
    SetIniValue IniData, "Config", "RemoveUnderscores", Abs(CLng(Me.RemoveUnderscores))
    SetIniValue IniData, "Config", "RemoveColons", Abs(CLng(Me.RemoveColons))
    SetIniValue IniData, "Config", "RemoveEmptyLines", Abs(CLng(Me.RemoveEmptyLines))
    SetIniValue IniData, "Config", "RemoveSpaces", Abs(CLng(Me.RemoveSpaces))
    
    SetIniValue IniData, "Config", "EncryptControlCaptions", Abs(CLng(Me.EncryptControlCaptions))
    SetIniValue IniData, "Config", "EncryptStrings", Abs(CLng(Me.EncryptStrings))
    SetIniValue IniData, "Config", "ExportStrings", Abs(CLng(Me.ExportStrings))
    SetIniValue IniData, "Config", "EnableCrackCheck", Abs(CLng(Me.EnableCrackCheck))
    
    SetIniValue IniData, "Config", "EnableDebugMode", Abs(CLng(Me.EnableDebugMode))
    SetIniValue IniData, "Config", "AddErrorHandlers", Abs(CLng(Me.AddErrorHandlers))
    
    If FileExist(IniPath) Then
        DeleteFile IniPath
    End If
    
    WriteFile IniPath, IniData
    
End Function

Function LoadConfig()
    
    If FileExist(IniPath) Then
        
        IniData = ReadFile(IniPath)
    
    Else
        
        Exit Function
        
    End If
    
    Me.Obfuscate = CBool(GetIniValue(IniData, "Config", "Obfuscate"))
    
    Me.ReplaceFileNames = CBool(GetIniValue(IniData, "Config", "ReplaceFileNames"))
    Me.ReplaceProjectNames = CBool(GetIniValue(IniData, "Config", "ReplaceProjectNames"))
    Me.ReplaceModuleNames = CBool(GetIniValue(IniData, "Config", "ReplaceModuleNames"))
    
    Me.RemoveCodeLayout = CBool(GetIniValue(IniData, "Config", "RemoveCodeLayout"))
    
    Me.RemoveComments = CBool(GetIniValue(IniData, "Config", "RemoveComments"))
    Me.RemoveUnderscores = CBool(GetIniValue(IniData, "Config", "RemoveUnderscores"))
    Me.RemoveColons = CBool(GetIniValue(IniData, "Config", "RemoveColons"))
    Me.RemoveEmptyLines = CBool(GetIniValue(IniData, "Config", "RemoveEmptyLines"))
    Me.RemoveSpaces = CBool(GetIniValue(IniData, "Config", "RemoveSpaces"))
    
    Me.EncryptControlCaptions = CBool(GetIniValue(IniData, "Config", "EncryptControlCaptions"))
    Me.EncryptStrings = CBool(GetIniValue(IniData, "Config", "EncryptStrings"))
    Me.ExportStrings = CBool(GetIniValue(IniData, "Config", "ExportStrings"))
    Me.EnableCrackCheck = CBool(GetIniValue(IniData, "Config", "EnableCrackCheck"))
    
    Me.EnableDebugMode = CBool(GetIniValue(IniData, "Config", "EnableDebugMode"))
    Me.AddErrorHandlers = CBool(GetIniValue(IniData, "Config", "AddErrorHandlers"))
   
    
End Function


