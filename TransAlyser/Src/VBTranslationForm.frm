VERSION 5.00
Begin VB.Form VBTranslationForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Language"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VBTranslationForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.ComboBox cbLanguage 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Please select your language:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "VBTranslationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Const SORT_DEFAULT As Integer = &H0
Private Const LANG_NEUTRAL As Integer = &H0
Private Const SUBLANG_DEFAULT As Integer = &H1
Private Const SUBLANG_SYS_DEFAULT As Integer = &H2

Private Const LANG_SYSTEM_DEFAULT As Long = (SUBLANG_SYS_DEFAULT * 1024&) Or LANG_NEUTRAL
Private Const LANG_USER_DEFAULT As Long = (SUBLANG_DEFAULT * 1024&) Or LANG_NEUTRAL

Private Const LOCALE_SYSTEM_DEFAULT As Long = (SORT_DEFAULT * 65536) Or LANG_SYSTEM_DEFAULT
Private Const LOCALE_USER_DEFAULT As Long = (SORT_DEFAULT * 65536) Or LANG_USER_DEFAULT

Private Const LOCALE_NOUSEROVERRIDE = &H80000000 '// do not use user overrides
Private Const LOCALE_USE_CP_ACP = &H40000000 '// use the system ACP
Private Const LOCALE_RETURN_NUMBER = &H20000000 '// return number instead of string

Private Const LOCALE_ILANGUAGE = &H1& '// language id
Private Const LOCALE_SLANGUAGE = &H2& '// localized name of language


Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    
    LoadLanguages
    
End Sub

Private Function GetLocaleLanguage() As String
On Error Resume Next
    
    Dim StrRet As String
    Dim x As Long

    StrRet = String$(1024, 0)

    If GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLANGUAGE, StrRet, Len(StrRet)) Then
        StrRet = Left$(StrRet, InStr(StrRet, vbNullChar) - 1)
        x = InStr(1, StrRet, " ", vbTextCompare)
        If x <> 0 Then
            StrRet = Left(StrRet, x - 1)
        End If
    End If
    
    GetLocaleLanguage = StrRet
    
End Function

Private Function FixPath(Path As String, Optional Terminate As Boolean = True) As String
On Error GoTo Error

    If Right(Path, 1) = "\" Then
        FixPath = Left(Path, Len(Path) - 1)
    Else
        FixPath = Path
    End If
    If Terminate Then FixPath = FixPath & "\"

Exit Function
Error:
    Resume Next
End Function

Private Function GetFileList(Path As String, FileArray() As String, Optional WithPath As Boolean = True) As Long
Dim FileName As String, sPath As String
Dim Files As Long
On Error GoTo Error
    sPath = FixPath(Path)
    FileName = Dir(sPath, vbArchive + vbHidden + vbReadOnly + vbSystem)
    Do
        Select Case FileName
            Case "", ".", ".."
            Case Else
                Files = Files + 1
                ReDim Preserve FileArray(Files)
                If WithPath Then
                    FileArray(Files) = sPath & FileName
                Else
                    FileArray(Files) = FileName
                End If
        End Select
        On Error Resume Next
        FileName = Dir()
        On Error GoTo Error
    Loop While FileName <> ""
    If Files = 0 Then ReDim FileArray(0)
    GetFileList = Files
Exit Function
Error:
    Resume Next
End Function

Private Function GetFileTitle(FileName As String)
    
    Dim x As Long
    Dim FileTitle As String
    
    x = InStrRev(FileName, ".")
    
    If x <> 0 Then
        FileTitle = Left(FileName, x - 1)
    Else
        FileTitle = FileName
    End If
    
    GetFileTitle = FileTitle
    
End Function

Private Sub LoadLanguages()
    
    Dim Files() As String
    
    Dim i As Long
    Dim c As Long
    
    Dim sLanguage As String
    
    c = GetFileList(App.Path, Files, False)
        
    For i = 1 To c
        
        If Right(Files(i), 4) = ".lng" Then
            
            sLanguage = GetFileTitle(Files(i))
            
            Me.cbLanguage.AddItem sLanguage
            
        End If
        
    Next i
    
    On Error Resume Next
    Me.cbLanguage.ListIndex = 0
    On Error GoTo 0
    
    For i = 0 To Me.cbLanguage.ListCount - 1
        
        If LCase(Me.cbLanguage.List(i)) = LCase(GetLocaleLanguage) Then
            
            Me.cbLanguage.ListIndex = i
            
            Exit For
            
        End If
        
    Next i
    
End Sub

Public Property Get SelectedLanguage() As String
    
    SelectedLanguage = Me.cbLanguage.Text
    
End Property

