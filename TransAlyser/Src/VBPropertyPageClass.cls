VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VBPropertyPageClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPath As String
Private pName As String

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

