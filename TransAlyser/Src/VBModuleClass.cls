VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBModuleClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPath As String
Private pName As String

Public ID As Long
Public TypeName As String
Public SrcPath As String 'Originale relative Pfad-Angabe in Ursprungsdatei
Public ResFileName As String 'gibt den Pfad zur Res Datei an 'z.B.: fMain.frx

Private pCode As String

Public Property Get Path() As String
    Path = pPath
End Property

Public Property Let Path(NewValue As String)
    pPath = NewValue
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(NewValue As String)
    pName = NewValue
End Property

Property Get FileName() As String
    FileName = ExtractFileName(Path)
End Property

Property Get FileTitle() As String
    FileTitle = GetFileTitle(Path)
End Property

Property Get FileExt() As String
    FileExt = TypeNameExt(TypeName)
End Property

Property Get NewFileName() As String
    NewFileName = Name & FileExt
End Property

Property Get NewResFileName() As String
    NewResFileName = Name & TypeNameResExt(TypeName)
End Property

Property Get RootDir() As String
    RootDir = ExtractDirName(Path)
End Property

Property Get Code() As String
    Code = pCode
End Property


