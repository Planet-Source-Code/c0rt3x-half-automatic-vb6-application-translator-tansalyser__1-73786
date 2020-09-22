VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyKeyWords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pWord() As String

Private Sub Class_Initialize()
    ReDim pWord(0)
End Sub

Public Property Get Count() As Long
    Count = UBound(pWord)
End Property

Public Property Get Word(Index As Long) As String
    Word = pWord(Index)
End Property


Public Function AddList(WordList As String)
    
    Dim WordList2 As String
    Dim NewWords() As String
    Dim i As Long
    
    WordList2 = Replace(WordList, vbNewLine, "")
    NewWords = Split(WordList2, " ")

    For i = 0 To UBound(NewWords)
        
        If Not Exists(NewWords(i)) Then
            
            ReDim Preserve pWord(UBound(pWord) + 1)
            
            pWord(UBound(pWord)) = NewWords(i)
        
        End If
        
    Next i
    
End Function

Public Function GetList() As String
    
    GetList = Join(pWord, " ")

End Function

Public Property Get Exists(ThisWord As String) As Boolean

    Dim i As Long
    
    For i = 1 To Count
        
        If LCase(pWord(i)) = LCase(ThisWord) Then
            Exists = True
            Exit Property
        End If
        
    Next i
    
End Property
