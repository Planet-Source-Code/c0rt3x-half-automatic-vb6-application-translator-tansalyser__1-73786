Attribute VB_Name = "MlStringModule"
Option Explicit

'API declarations
Public Const LOCALE_SLANGUAGE       As Long = &H2  'localized name of Language
Public Const LOCALE_SABBREVLANGNAME As Long = &H3  'abbreviated language name
Public Const LOCALE_SNATIVELANGNAME As Long = &H4  'native name of Language

Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" _
                        (ByVal Locale As Long, _
                         ByVal LCType As Long, _
                         ByVal lpLCData As String, _
                         ByVal cchData As Long) As Long

'Variables
Public ml_CurrentLanguageId      As Long
Public Const ml_LanguageCount    As Long = 2
Public Const ml_OriginalLanguage As Long = 2057

'Functions
Public Function ml_string(ByVal StringID As Long, Optional ByVal Text As String = "") As String
  ml_string = Text
End Function

Public Function ml_LanguageName(ByVal LangIndex As Long) As String
  ml_LanguageName = "Invalid Language Index"
End Function

Public Sub ml_ChangeLanguage(ByVal LanguageID As Long, ByVal Language As String)
  
  Dim LanguageIDs As Variant
  Dim Index       As Long
  
  'This function may be called from the ml_RuntimeSupport_LanguageChanged event.
  'This event is used to change the language across separately compiled components
  '(exe, dll, ocx). In this case, the components should support the same languages
  'and use the same IDs. Using non standard language IDs is not recommended.
  
  'The following loop checks that the specified language is supported by this
  'component. If not, then the original language is used.
  
  LanguageIDs = ml_LanguageIds
  ml_CurrentLanguageId = ml_OriginalLanguage
  For Index = LBound(LanguageIDs) To UBound(LanguageIDs)
    If LanguageIDs(Index) = LanguageID Then
      ml_CurrentLanguageId = LanguageID
      Exit For
    End If
  Next
  
End Sub

Public Function ml_LanguageIds() As Variant
  ml_LanguageIds = Array(2057, 1031)
End Function

'Helper function for using GetLocaleIndo
Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, _
                                  ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim nSize As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   nSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
   If nSize Then
    
     'pad a buffer with spaces
      sReturn = Space$(nSize)
       
     'and call again passing the buffer
      nSize = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (nSize > 0)
      If nSize Then
      
        'nSize holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, nSize - 1)
      
      End If
   
   End If
    
End Function

'Alternative to ml_LanguageName.
Public Function ml_LocaleName(ByVal LangIndex As Long) As String
  ml_LocaleName = GetUserLocaleInfo(LangIndex, LOCALE_SNATIVELANGNAME)
End Function

'Pathetic function for text substitution.
'Use place holders in the form %0, %1, %2 ...
'Example use:
'MsgBox Substitute("Do you want to delete the file %0",FileName)
Public Function Substitute(ByVal FormatString As String, ParamArray Params() As Variant) As String
  Dim Index As Long
  For Index = LBound(Params) To UBound(Params)
    FormatString = Replace(FormatString, "%" & Index, Params(Index))
  Next
  Substitute = FormatString
End Function


