Attribute VB_Name = "Common_NetworkModule"
Private Declare Function IsDestinationReachable Lib _
  "Sensapi.dll" Alias "IsDestinationReachableA" _
  (ByVal lpszDestination As String, _
  lpQOCInfo As QOCINFO) As Long

Private Type QOCINFO
  dwSize As Long
  dwFlags As Long
  dwInSpeed As Long
  dwOutSpeed As Long
End Type

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal lpszAgent As String, ByVal dwAccessType As Long, ByVal lpszProxyName As String, ByVal lpszProxyBypass As String, ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInet As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Long

Public Function Ping(ByVal IP As String) As Boolean
On Error GoTo Error
    
    Dim QuestStruct As QOCINFO
    Dim lReturn As Long

    ' Größe der Struktur
    QuestStruct.dwSize = Len(QuestStruct)

    ' Prüfen, ob Ziel erreichbar
    lReturn = IsDestinationReachable(IP, QuestStruct)
  
    ' Antwort auswerten
    If lReturn = 1 Then
        ' Antwort bekommen
        Ping = True
    Else
        ' keine Antwort
        Ping = False
    End If

Exit Function
Error:
    Assert , "NetworkModule.Ping", Err.Number, Err.Description, "IP: '" & CStr(IP) & "'"
    Resume Next
End Function

Function IsOnline() As Boolean
On Error Resume Next
   Dim hInet As Long
   Dim hUrl As Long
   Dim Flags As Long
   Dim url As Variant
   hInet = InternetOpen(App.Title, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0&)
   If hInet Then
      Flags = INTERNET_FLAG_KEEP_CONNECTION Or INTERNET_FLAG_NO_CACHE_WRITE Or INTERNET_FLAG_RELOAD
      hUrl = InternetOpenUrl(hInet, "http://www.yahoo.com", vbNullString, 0, Flags, 0)
      If hUrl Then IsOnline = True
   End If
   InternetCloseHandle hInet
End Function
