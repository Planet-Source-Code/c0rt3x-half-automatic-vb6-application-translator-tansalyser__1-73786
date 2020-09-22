VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBFormClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPath As String
Private pName As String

Public TypeName As String
Public SrcPath As String 'Originale relative Pfad-Angabe in Ursprungsdatei
Public ResFileName As String 'gibt den Pfad zur Res Datei an 'z.B.: fMain.frx

Public Head As String
Public Body As String

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


Function Load(Path As String)
Dim Bin As String
Dim Line() As String
Dim c As String
Dim i As Long
Dim ValueName As String
Dim Value As String
Dim s As String

    pPath = Path
    Bin = ReadFile(Path)
    Line = Split(Bin, vbNewLine)
    
    For i = 0 To UBound(Line)
        If Left(Line(i), Len("Attribute VB_Name = ")) = "Attribute VB_Name = " Then
            Name = FilterString(Mid(Line(i), Len("Attribute VB_Name = ") + 1), Chr(34))
        End If
        
        If InStr(1, Line(i), TypeNameResExt(TypeName) & Chr(34) & ":") Then
            If TypeNameResExt(TypeName) <> "" Then
                s = Mid(Line(i), InStr(1, Line(i), "=") + 1)
                s = LTrim(s)
                s = Left(s, InStr(2, s, Chr(34)) - 1)
                s = Mid(s, 2)
                ResFileName = s
            End If
        End If
    Next i
    
End Function


