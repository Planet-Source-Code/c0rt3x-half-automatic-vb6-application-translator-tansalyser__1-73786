Attribute VB_Name = "mKeyPressed"
Option Explicit

' *** Den Artikel zu diesem Modul finden Sie unter http://www.aboutvb.de/khw/artikel/khwkeypressed.htm ***

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Enum KeyCodeExConstants
' leichter merkbar
    vbKeyAlt = vbKeyMenu
    vbKeyPrintScreen = vbKeySnapshot
    
' Windows-Tasten
    vbKeyLWin = &H5B 'Linke Windows-Taste
    vbKeyRWin = &H5C 'Rechte Windows-Taste
    vbKeyApps = &H5D 'Anwendungen-Taste
    
' Nur unter NT/Windows 2000 und später
    vbKeyLShift = &HA0
    vbKeyRShift = &HA1
    vbKeyLControl = &HA2
    vbKeyRControl = &HA3
    vbKeyLAlt = &HA4
    vbKeyRAlt = &HA5
    ' bzw. im Original
    vbKeyLMenu = &HA4
    vbKeyRMenu = &HA5
    
' ab Windows 2000
    vbKeyXButton1 = &H5 'X1-Maus-Taste
    vbKeyXButton2 = &H6 'X2-Maus-Taste
    vbKeyBrowserBack = &HA6
    vbKeyBrowserForward = &HA7
    vbKeyBrowserRefresh = &HA8
    vbKeyBrowserStop = &HA9
    vbKeyBrowserSearch = &HAA
    vbKeyBrowserFavorites = &HAB
    vbKeyBrowserHome = &HAC
    vbKeyVolumeMute = &HAD
    vbKeyVolumeDown = &HAE
    vbKeyVolumeUp = &HAF
    vbKeyMediaNextTrack = &HB0
    vbKeyMediaPrevTrack = &HB1
    vbKeyMediaStop = &HB2
    vbKeyMediaPlayPause = &HB3
    vbKeyLaunchMail = &HB4
    vbKeyLaunchMediaSelect = &HB5
    vbKeyLaunchApp1 = &HB6
    vbKeyLaunchApp2 = &HB7
    
' IME
    vbKeyKANA = &H15
    vbKeyHANGUL = &H15
    vbKeyJUNJA = &H17
    vbKeyFINAL = &H18
    vbKeyHANJA = &H19
    vbKeyKANJI = &H19
    vbKeyConvert = &H1C
    vbKeyNonConvert = &H1D
    vbKeyAccept = &H1E
    vbKeyModeChange = &H1F
    vbKeyProcessKey = &HE5
    
' weitere Funktionstasten
    vbKeyF17 = &H80
    vbKeyF18 = &H81
    vbKeyF19 = &H82
    vbKeyF20 = &H83
    vbKeyF21 = &H84
    vbKeyF22 = &H85
    vbKeyF23 = &H86
    vbKeyF24 = &H87
    
' OEM-Zuordnungen, ab Windows 2000 - international
    vbKeyOEMPlus = &HBB
    vbKeyOEMComma = &HBC
    vbKeyOEMMinus = &HBD
    vbKeyOEMPeriod = &HBE
    
' OEM-Zuordnungen, ab Windows 2000 - US-Keyboards
    vbKeyOEM_1 = &HBA ' Taste ;:
    vbKeyOEM_2 = &HBF ' Taste /?
    vbKeyOEM_3 = &HC0 ' Taste `~
    vbKeyOEM_4 = &HDB ' Taste [{
    vbKeyOEM_5 = &HDC ' Taste \|
    vbKeyOEM_6 = &HDD ' Taste ]}
    vbKeyOEM_7 = &HDE ' Taste '"
    
' Sonstige Sondertasten
    vbKeySleep = &H5F
    vbKeyATTN = &HF6
    vbKeyCRSEL = &HF7
    vbKeyEXSEL = &HF8
    vbKeyEREOF = &HF9
    vbKeyPlay = &HFA
    vbKeyZoom = &HFB
    vbKeyPA1 = &HFD
    vbKeyOEMClear = &HFE
End Enum

Public Function KeyPressed(ByVal Key As Long) As Boolean
    KeyPressed = CBool((GetAsyncKeyState(Key) And &H8000) = &H8000)
End Function

Public Function AltGrPressed() As Boolean
    AltGrPressed = CBool((GetAsyncKeyState(vbKeyControl) And &H8000) = &H8000) And CBool((GetAsyncKeyState(vbKeyAlt) And &H8000) = &H8000)
End Function

Public Function MouseButtonPressed(ByVal MouseKey As Long, Optional ByVal Logical As Boolean) As Boolean
    Dim nMouseKey As Long
    
    Const SM_SWAPBUTTON = 23
    
    Select Case MouseKey
        Case vbKeyLButton, vbKeyRButton, vbKeyMButton
            If Logical Then
                If GetSystemMetrics(SM_SWAPBUTTON) Then
                    Select Case MouseKey
                        Case vbKeyLButton
                            nMouseKey = vbKeyRButton
                        Case vbKeyRButton
                            nMouseKey = vbKeyLButton
                        Case vbKeyMButton
                            nMouseKey = MouseKey
                    End Select
                Else
                    nMouseKey = MouseKey
                End If
            Else
                nMouseKey = MouseKey
            End If
        Case Else
            Err.Raise 380
    End Select
    MouseButtonPressed = CBool((GetAsyncKeyState(nMouseKey) And &H8000) = &H8000)
End Function



