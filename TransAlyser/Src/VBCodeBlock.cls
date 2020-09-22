VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBCodeBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ParentModule     As VBModule
Public ParentBlock      As VBCodeBlock

Public ID               As Long
Public Name             As String

Public Code             As String
Public Protected        As Boolean

'Public StartPos         As Long
'Public EndPos           As Long
'Public Length           As Long

Public FirstLine        As Long
Public LastLine         As Long
Public LineCount        As Long

Private pChildCount     As Long
Private pChilds()       As VBCodeBlock

Private Sub Class_Initialize()
    ReDim pChilds(0)
End Sub

Property Get ChildCount() As Long
    ChildCount = pChildCount
End Property

Property Let ChildCount(NewValue As Long)
    ReDim Preserve pChilds(NewValue)
    pChildCount = NewValue
End Property

Property Get Childs(Index As Long) As VBCodeBlock
    Set Childs = pChilds(Index)
End Property

Property Set Childs(Index As Long, NewObject As VBCodeBlock)
    Set pChilds(Index) = NewObject
End Property
