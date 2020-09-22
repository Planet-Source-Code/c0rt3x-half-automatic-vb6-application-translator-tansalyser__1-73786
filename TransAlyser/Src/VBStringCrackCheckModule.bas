Attribute VB_Name = "VBStringCrackCheckModule"
Option Explicit

Private Const VBStringBase64Key As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789%&"


Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Const GCL_HMODULE = -16


Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Function VBLoadString(VBStringID As Long) As String
    
    Dim FileData As String
    Dim Lines() As String
    Dim LineStr As String
    
    FileData = RC4(VBLoadFile(App.ExeName & ".res"), GetAppFingerPrint())
    
    Lines = Split(FileData, vbNewLine)
    LineStr = Lines(VBStringID - 1)
    
    VBLoadString = LineStr
    
End Function

Private Function VBStringBase64Decode(Text As String) As String
On Error GoTo Error

    Dim w1 As Integer
    Dim w2 As Integer
    Dim w3 As Integer
    Dim w4 As Integer
    Dim n As Integer
    Dim retry As String
    
    For n = 1 To Len(Text) Step 4
        w1 = VBStringMimeDecode(Mid$(Text, n, 1))
        w2 = VBStringMimeDecode(Mid$(Text, n + 1, 1))
        w3 = VBStringMimeDecode(Mid$(Text, n + 2, 1))
        w4 = VBStringMimeDecode(Mid$(Text, n + 3, 1))
        If w2 >= 0 Then retry = retry + Chr$(((w1 * 4 + Int(w2 / 16)) And 255))
        If w3 >= 0 Then retry = retry + Chr$(((w2 * 16 + Int(w3 / 4)) And 255))
        If w4 >= 0 Then retry = retry + Chr$(((w3 * 64 + w4) And 255))
    Next
    
    VBStringBase64Decode = retry

Exit Function
Error:
    Debug.Print Err.Description
    Resume Next
End Function


Private Function VBStringMimeDecode(a As String) As Integer
On Error GoTo Error

    If Len(a) = 0 Then
        VBStringMimeDecode = -1
        Exit Function
    End If
    VBStringMimeDecode = InStr(VBStringBase64Key, a) - 1

Exit Function
Error:
    Debug.Print Err.Description
    Resume Next
End Function


Private Function GetAppEXEPath(Optional hwnd As Long) As String
On Error GoTo Error
    
    Dim nResult As String
    Dim nLenResult As Long
    
    nResult = Space(260)
    nLenResult = Len(nResult)
    nLenResult = GetModuleFileName(GetClassLong(hwnd, GCL_HMODULE), _
    nResult, nLenResult)
    GetAppEXEPath = Left$(nResult, nLenResult)
    
Exit Function
Error:
    
End Function

Function VBLoadFile(FilePath As String, Optional Start As Long, Optional Lenght As Long, Optional Reverse As Boolean = False) As String
On Error GoTo Error

    Dim FileNum As Long
    Dim Buffer As String
    Dim lLen As Long
    Dim lStart As Long
    Dim lLOF As Long

    lLen = Lenght
    lStart = Start
    FileNum = FreeFile
    Open FilePath For Binary As #FileNum
        lLOF = LOF(FileNum)
        If lStart > lLOF Then
            Err.Description = "Startpoint > Filelenght"
            GoTo Error
        End If
        If lStart = 0 Then lStart = 1
        If lLen = 0 Or (lStart + lLen - 1) > lLOF Then lLen = (lLOF - lStart + 1)
        If Reverse Then lStart = lLOF - lStart - lLen + 2
        Buffer = String(lLen, " ")
        Get #FileNum, lStart, Buffer
    Close FileNum
    VBLoadFile = Buffer
    
Exit Function
Error:
    MsgBox Err.Description
    Resume Next
End Function

Private Function CalcCRC32(ByteArray() As Byte) As Long
    
    Dim m_CRC32 As Long
    Dim m_CRC32Asm() As Byte
    Dim m_CRC32Table(0 To 255) As Long
    
    Dim i As Long
    Dim sASM As String
  
    Dim ByteSize As Long
    
    m_CRC32Table(0) = &H0
    m_CRC32Table(1) = &H77073096
    m_CRC32Table(2) = &HEE0E612C
    m_CRC32Table(3) = &H990951BA
    m_CRC32Table(4) = &H76DC419
    m_CRC32Table(5) = &H706AF48F
    m_CRC32Table(6) = &HE963A535
    m_CRC32Table(7) = &H9E6495A3
    m_CRC32Table(8) = &HEDB8832
    m_CRC32Table(9) = &H79DCB8A4
    m_CRC32Table(10) = &HE0D5E91E
    m_CRC32Table(11) = &H97D2D988
    m_CRC32Table(12) = &H9B64C2B
    m_CRC32Table(13) = &H7EB17CBD
    m_CRC32Table(14) = &HE7B82D07
    m_CRC32Table(15) = &H90BF1D91
    m_CRC32Table(16) = &H1DB71064
    m_CRC32Table(17) = &H6AB020F2
    m_CRC32Table(18) = &HF3B97148
    m_CRC32Table(19) = &H84BE41DE
    m_CRC32Table(20) = &H1ADAD47D
    m_CRC32Table(21) = &H6DDDE4EB
    m_CRC32Table(22) = &HF4D4B551
    m_CRC32Table(23) = &H83D385C7
    m_CRC32Table(24) = &H136C9856
    m_CRC32Table(25) = &H646BA8C0
    m_CRC32Table(26) = &HFD62F97A
    m_CRC32Table(27) = &H8A65C9EC
    m_CRC32Table(28) = &H14015C4F
    m_CRC32Table(29) = &H63066CD9
    m_CRC32Table(30) = &HFA0F3D63
    m_CRC32Table(31) = &H8D080DF5
    m_CRC32Table(32) = &H3B6E20C8
    m_CRC32Table(33) = &H4C69105E
    m_CRC32Table(34) = &HD56041E4
    m_CRC32Table(35) = &HA2677172
    m_CRC32Table(36) = &H3C03E4D1
    m_CRC32Table(37) = &H4B04D447
    m_CRC32Table(38) = &HD20D85FD
    m_CRC32Table(39) = &HA50AB56B
    m_CRC32Table(40) = &H35B5A8FA
    m_CRC32Table(41) = &H42B2986C
    m_CRC32Table(42) = &HDBBBC9D6
    m_CRC32Table(43) = &HACBCF940
    m_CRC32Table(44) = &H32D86CE3
    m_CRC32Table(45) = &H45DF5C75
    m_CRC32Table(46) = &HDCD60DCF
    m_CRC32Table(47) = &HABD13D59
    m_CRC32Table(48) = &H26D930AC
    m_CRC32Table(49) = &H51DE003A
    m_CRC32Table(50) = &HC8D75180
    m_CRC32Table(51) = &HBFD06116
    m_CRC32Table(52) = &H21B4F4B5
    m_CRC32Table(53) = &H56B3C423
    m_CRC32Table(54) = &HCFBA9599
    m_CRC32Table(55) = &HB8BDA50F
    m_CRC32Table(56) = &H2802B89E
    m_CRC32Table(57) = &H5F058808
    m_CRC32Table(58) = &HC60CD9B2
    m_CRC32Table(59) = &HB10BE924
    m_CRC32Table(60) = &H2F6F7C87
    m_CRC32Table(61) = &H58684C11
    m_CRC32Table(62) = &HC1611DAB
    m_CRC32Table(63) = &HB6662D3D
    m_CRC32Table(64) = &H76DC4190
    m_CRC32Table(65) = &H1DB7106
    m_CRC32Table(66) = &H98D220BC
    m_CRC32Table(67) = &HEFD5102A
    m_CRC32Table(68) = &H71B18589
    m_CRC32Table(69) = &H6B6B51F
    m_CRC32Table(70) = &H9FBFE4A5
    m_CRC32Table(71) = &HE8B8D433
    m_CRC32Table(72) = &H7807C9A2
    m_CRC32Table(73) = &HF00F934
    m_CRC32Table(74) = &H9609A88E
    m_CRC32Table(75) = &HE10E9818
    m_CRC32Table(76) = &H7F6A0DBB
    m_CRC32Table(77) = &H86D3D2D
    m_CRC32Table(78) = &H91646C97
    m_CRC32Table(79) = &HE6635C01
    m_CRC32Table(80) = &H6B6B51F4
    m_CRC32Table(81) = &H1C6C6162
    m_CRC32Table(82) = &H856530D8
    m_CRC32Table(83) = &HF262004E
    m_CRC32Table(84) = &H6C0695ED
    m_CRC32Table(85) = &H1B01A57B
    m_CRC32Table(86) = &H8208F4C1
    m_CRC32Table(87) = &HF50FC457
    m_CRC32Table(88) = &H65B0D9C6
    m_CRC32Table(89) = &H12B7E950
    m_CRC32Table(90) = &H8BBEB8EA
    m_CRC32Table(91) = &HFCB9887C
    m_CRC32Table(92) = &H62DD1DDF
    m_CRC32Table(93) = &H15DA2D49
    m_CRC32Table(94) = &H8CD37CF3
    m_CRC32Table(95) = &HFBD44C65
    m_CRC32Table(96) = &H4DB26158
    m_CRC32Table(97) = &H3AB551CE
    m_CRC32Table(98) = &HA3BC0074
    m_CRC32Table(99) = &HD4BB30E2
    m_CRC32Table(100) = &H4ADFA541
    m_CRC32Table(101) = &H3DD895D7
    m_CRC32Table(102) = &HA4D1C46D
    m_CRC32Table(103) = &HD3D6F4FB
    m_CRC32Table(104) = &H4369E96A
    m_CRC32Table(105) = &H346ED9FC
    m_CRC32Table(106) = &HAD678846
    m_CRC32Table(107) = &HDA60B8D0
    m_CRC32Table(108) = &H44042D73
    m_CRC32Table(109) = &H33031DE5
    m_CRC32Table(110) = &HAA0A4C5F
    m_CRC32Table(111) = &HDD0D7CC9
    m_CRC32Table(112) = &H5005713C
    m_CRC32Table(113) = &H270241AA
    m_CRC32Table(114) = &HBE0B1010
    m_CRC32Table(115) = &HC90C2086
    m_CRC32Table(116) = &H5768B525
    m_CRC32Table(117) = &H206F85B3
    m_CRC32Table(118) = &HB966D409
    m_CRC32Table(119) = &HCE61E49F
    m_CRC32Table(120) = &H5EDEF90E
    m_CRC32Table(121) = &H29D9C998
    m_CRC32Table(122) = &HB0D09822
    m_CRC32Table(123) = &HC7D7A8B4
    m_CRC32Table(124) = &H59B33D17
    m_CRC32Table(125) = &H2EB40D81
    m_CRC32Table(126) = &HB7BD5C3B
    m_CRC32Table(127) = &HC0BA6CAD
    m_CRC32Table(128) = &HEDB88320
    m_CRC32Table(129) = &H9ABFB3B6
    m_CRC32Table(130) = &H3B6E20C
    m_CRC32Table(131) = &H74B1D29A
    m_CRC32Table(132) = &HEAD54739
    m_CRC32Table(133) = &H9DD277AF
    m_CRC32Table(134) = &H4DB2615
    m_CRC32Table(135) = &H73DC1683
    m_CRC32Table(136) = &HE3630B12
    m_CRC32Table(137) = &H94643B84
    m_CRC32Table(138) = &HD6D6A3E
    m_CRC32Table(139) = &H7A6A5AA8
    m_CRC32Table(140) = &HE40ECF0B
    m_CRC32Table(141) = &H9309FF9D
    m_CRC32Table(142) = &HA00AE27
    m_CRC32Table(143) = &H7D079EB1
    m_CRC32Table(144) = &HF00F9344
    m_CRC32Table(145) = &H8708A3D2
    m_CRC32Table(146) = &H1E01F268
    m_CRC32Table(147) = &H6906C2FE
    m_CRC32Table(148) = &HF762575D
    m_CRC32Table(149) = &H806567CB
    m_CRC32Table(150) = &H196C3671
    m_CRC32Table(151) = &H6E6B06E7
    m_CRC32Table(152) = &HFED41B76
    m_CRC32Table(153) = &H89D32BE0
    m_CRC32Table(154) = &H10DA7A5A
    m_CRC32Table(155) = &H67DD4ACC
    m_CRC32Table(156) = &HF9B9DF6F
    m_CRC32Table(157) = &H8EBEEFF9
    m_CRC32Table(158) = &H17B7BE43
    m_CRC32Table(159) = &H60B08ED5
    m_CRC32Table(160) = &HD6D6A3E8
    m_CRC32Table(161) = &HA1D1937E
    m_CRC32Table(162) = &H38D8C2C4
    m_CRC32Table(163) = &H4FDFF252
    m_CRC32Table(164) = &HD1BB67F1
    m_CRC32Table(165) = &HA6BC5767
    m_CRC32Table(166) = &H3FB506DD
    m_CRC32Table(167) = &H48B2364B
    m_CRC32Table(168) = &HD80D2BDA
    m_CRC32Table(169) = &HAF0A1B4C
    m_CRC32Table(170) = &H36034AF6
    m_CRC32Table(171) = &H41047A60
    m_CRC32Table(172) = &HDF60EFC3
    m_CRC32Table(173) = &HA867DF55
    m_CRC32Table(174) = &H316E8EEF
    m_CRC32Table(175) = &H4669BE79
    m_CRC32Table(176) = &HCB61B38C
    m_CRC32Table(177) = &HBC66831A
    m_CRC32Table(178) = &H256FD2A0
    m_CRC32Table(179) = &H5268E236
    m_CRC32Table(180) = &HCC0C7795
    m_CRC32Table(181) = &HBB0B4703
    m_CRC32Table(182) = &H220216B9
    m_CRC32Table(183) = &H5505262F
    m_CRC32Table(184) = &HC5BA3BBE
    m_CRC32Table(185) = &HB2BD0B28
    m_CRC32Table(186) = &H2BB45A92
    m_CRC32Table(187) = &H5CB36A04
    m_CRC32Table(188) = &HC2D7FFA7
    m_CRC32Table(189) = &HB5D0CF31
    m_CRC32Table(190) = &H2CD99E8B
    m_CRC32Table(191) = &H5BDEAE1D
    m_CRC32Table(192) = &H9B64C2B0
    m_CRC32Table(193) = &HEC63F226
    m_CRC32Table(194) = &H756AA39C
    m_CRC32Table(195) = &H26D930A
    m_CRC32Table(196) = &H9C0906A9
    m_CRC32Table(197) = &HEB0E363F
    m_CRC32Table(198) = &H72076785
    m_CRC32Table(199) = &H5005713
    m_CRC32Table(200) = &H95BF4A82
    m_CRC32Table(201) = &HE2B87A14
    m_CRC32Table(202) = &H7BB12BAE
    m_CRC32Table(203) = &HCB61B38
    m_CRC32Table(204) = &H92D28E9B
    m_CRC32Table(205) = &HE5D5BE0D
    m_CRC32Table(206) = &H7CDCEFB7
    m_CRC32Table(207) = &HBDBDF21
    m_CRC32Table(208) = &H86D3D2D4
    m_CRC32Table(209) = &HF1D4E242
    m_CRC32Table(210) = &H68DDB3F8
    m_CRC32Table(211) = &H1FDA836E
    m_CRC32Table(212) = &H81BE16CD
    m_CRC32Table(213) = &HF6B9265B
    m_CRC32Table(214) = &H6FB077E1
    m_CRC32Table(215) = &H18B74777
    m_CRC32Table(216) = &H88085AE6
    m_CRC32Table(217) = &HFF0F6A70
    m_CRC32Table(218) = &H66063BCA
    m_CRC32Table(219) = &H11010B5C
    m_CRC32Table(220) = &H8F659EFF
    m_CRC32Table(221) = &HF862AE69
    m_CRC32Table(222) = &H616BFFD3
    m_CRC32Table(223) = &H166CCF45
    m_CRC32Table(224) = &HA00AE278
    m_CRC32Table(225) = &HD70DD2EE
    m_CRC32Table(226) = &H4E048354
    m_CRC32Table(227) = &H3903B3C2
    m_CRC32Table(228) = &HA7672661
    m_CRC32Table(229) = &HD06016F7
    m_CRC32Table(230) = &H4969474D
    m_CRC32Table(231) = &H3E6E77DB
    m_CRC32Table(232) = &HAED16A4A
    m_CRC32Table(233) = &HD9D65ADC
    m_CRC32Table(234) = &H40DF0B66
    m_CRC32Table(235) = &H37D83BF0
    m_CRC32Table(236) = &HA9BCAE53
    m_CRC32Table(237) = &HDEBB9EC5
    m_CRC32Table(238) = &H47B2CF7F
    m_CRC32Table(239) = &H30B5FFE9
    m_CRC32Table(240) = &HBDBDF21C
    m_CRC32Table(241) = &HCABAC28A
    m_CRC32Table(242) = &H53B39330
    m_CRC32Table(243) = &H24B4A3A6
    m_CRC32Table(244) = &HBAD03605
    m_CRC32Table(245) = &HCDD70693
    m_CRC32Table(246) = &H54DE5729
    m_CRC32Table(247) = &H23D967BF
    m_CRC32Table(248) = &HB3667A2E
    m_CRC32Table(249) = &HC4614AB8
    m_CRC32Table(250) = &H5D681B02
    m_CRC32Table(251) = &H2A6F2B94
    m_CRC32Table(252) = &HB40BBE37
    m_CRC32Table(253) = &HC30C8EA1
    m_CRC32Table(254) = &H5A05DF1B
    m_CRC32Table(255) = &H2D02EF8D
    
    sASM = "5589E557565053518B45088B008B750C8B7D108B4D1431DB8A1E30C3C1E80833049F464975F28B4D088901595B585E5F89EC5DC21000"
    
    ReDim m_CRC32Asm(0 To Len(sASM) \ 2 - 1)
    
    For i = 1 To Len(sASM) Step 2
        m_CRC32Asm(i \ 2) = Val("&H" & Mid$(sASM, i, 2))
    Next
    
    On Local Error GoTo NoData
    
    ByteSize = UBound(ByteArray) - LBound(ByteArray) + 1
    
    On Local Error GoTo 0
   
    Call CallWindowProc(VarPtr(m_CRC32Asm(0)), VarPtr(m_CRC32), VarPtr(ByteArray(LBound(ByteArray))), VarPtr(m_CRC32Table(0)), ByteSize)
    
NoData:

    CalcCRC32 = (Not m_CRC32)
    
End Function

Private Function GetCRC32FromFile(FilePath As String) As Long

    Dim Filenr As Integer
    Dim ByteArray() As Byte
    
    If (FileLen(FilePath) = 0) Then Exit Function
    
    Filenr = FreeFile
    Open FilePath For Binary As #Filenr
    ReDim ByteArray(0 To LOF(Filenr) - 1)
    Get #Filenr, , ByteArray()
    Close #Filenr
    
    GetCRC32FromFile = CalcCRC32(ByteArray)

End Function

Public Function GetAppFingerPrint() As String
    GetAppFingerPrint = CStr(GetCRC32FromFile(GetAppEXEPath))
End Function

Private Function RC4(sString As String, Optional sPassword As String) As String
On Error GoTo Error
    
    Dim s%(0 To 255), kep%(0 To 255)
    Dim a&, b%, i%, j%, K%, temp%
    Dim pwd$
    Dim cipherby As Byte, cipher$
    
    pwd = CStr(Len(sString)) & CStr(Len(sPassword)) & sPassword
    b = 0
    For a = 0 To 255
        b = b + 1
        If b > Len(pwd) Then
            b = 1
        End If
        kep(a) = Asc(Mid$(pwd, b, 1))
    Next a
    For a = 0 To 255
        s(a) = a
    Next a
    b = 0
    For a = 0 To 255
        b = (b + s(a) + kep(a)) Mod 256
        temp = s(a)
        s(a) = s(b)
        s(b) = temp
    Next a
    For a = 1 To Len(sString)
        i = (i + 1) Mod 256
        j = (j + s(i)) Mod 256
        temp = s(i)
        s(i) = s(j)
        s(j) = temp
        K = s((s(i) + s(j)) Mod 256)
        cipherby = Asc(Mid$(sString, a, 1)) Xor K
        cipher = cipher & Chr(cipherby)
    Next a
    
    RC4 = cipher

Exit Function
Error:
    Debug.Print Err.Description
End Function
