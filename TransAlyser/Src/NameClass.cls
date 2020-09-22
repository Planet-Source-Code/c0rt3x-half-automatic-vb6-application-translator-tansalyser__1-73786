VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NameClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pOldNames()     As String
Private pNameTypes()    As Long
Private pNewNames()     As String

Private Const NewLen    As Integer = 5

Property Get Count() As Long
    Count = UBound(pOldNames)
End Property

Property Get OldNames(Index As Long) As String
    OldNames = pOldNames(Index)
End Property

Property Get NewNames(Index As Long) As String
    NewNames = pNewNames(Index)
End Property

Property Get Exist(Name As String, Optional SearchForNewNames As Boolean) As Boolean
    Dim i As Long
    For i = 1 To Count
        If pOldNames(i) = Name Then
            Exist = True
            Exit Property
        End If
        If SearchForNewNames Then
            If pNewNames(i) = Name Then
                Exist = True
                Exit Property
            End If
        End If
    Next i
End Property

Function Add(Name As String, Optional Ext As String, Optional VBType As VBNameTypes) As String
    
    Dim RndName As String
    Dim NewName As String
    Dim c As Long
    Dim x As Long
    
    If Exist(Name) Then
        Add = Me.NewName(Name)
        Exit Function
    End If
    
    Randomize Timer
    
    x = RandomNumber(2, NewLen)
    RndName = RandomName(NewLen - 1)
    RndName = Left(RndName, x - 1) & RandomNumber(0, 9) & Mid(RndName, x)
    NewName = RndName
    Do While Exist(NewName)
        c = c + 1
        NewName = RndName & c
    Loop
    ReDim Preserve pOldNames(UBound(pOldNames) + 1)
    ReDim Preserve pNameTypes(UBound(pNameTypes) + 1)
    ReDim Preserve pNewNames(UBound(pNewNames) + 1)
    pOldNames(Count) = Name
    pNewNames(Count) = NewName & Ext
    Add = NewName
End Function

Property Get NewName(OldName As String)
    Dim i As Long
    For i = 1 To Count
        If pOldNames(i) = OldName Then
            NewName = pNewNames(i)
        End If
    Next i
End Property

Property Get NameType(OldName As String)
    Dim i As Long
    For i = 1 To Count
        If pOldNames(i) = OldName Then
            NameType = pNameTypes(i)
        End If
    Next i
End Property

Private Sub Class_Initialize()
    ReDim pOldNames(0)
    ReDim pNameTypes(0)
    ReDim pNewNames(0)
End Sub
