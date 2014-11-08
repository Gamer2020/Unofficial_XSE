Attribute VB_Name = "modTextSapp"
Option Explicit

Private AsciiReady As Boolean
Private SappReady As Boolean
Private BrailleReady As Boolean
Private bAsc2Sapp() As Byte
Private bAsc2Braille() As Byte
Private sSapp2Asc() As String
Private sSapp2AscJap() As String
Private sBraille2Asc() As String
Private bEnd(0) As Byte
Private cString As cStringBuilder

Private Declare Sub RtlMoveMemory Lib "kernel32" (pDest As Any, pSource As Any, ByVal cb As Long)

Public Function IsHex(ByRef sHex As String) As Boolean
Dim lTemp As Long
    
    On Error GoTo NotHex
    
    If LenB(sHex) <> 0 Then
        lTemp = CLng("&H" & sHex)
        IsHex = True
    End If
    
    Exit Function
    
NotHex:
End Function

Private Function IsHexI(ByRef iNibble1 As Integer, ByVal iNibble2 As Integer) As Boolean
    
    Select Case iNibble1
        Case vbKey0 To vbKey9, vbKeyA To vbKeyF, vbKeyA + 32 To vbKeyF + 32
            Select Case iNibble2
                Case vbKey0 To vbKey9, vbKeyA To vbKeyF, vbKeyA + 32 To vbKeyF + 32
                    IsHexI = True
            End Select
    End Select

End Function

Public Function PadHex$(ByVal lNumber As Long, lSize As Long)
    PadHex$ = RightB$("0000000" & Hex$(lNumber), lSize * 2&)
End Function

Private Sub AscInitialize()

    ReDim bAsc2Sapp(&HFF) As Byte
    ReDim bAsc2Braille(&HFF) As Byte
        
    bAsc2Sapp(&H20) = &H0 ' " "
    bAsc2Sapp(&HC0) = &H1 ' "À"
    bAsc2Sapp(&HC1) = &H2 ' "Á"
    bAsc2Sapp(&HC2) = &H3 ' "Â"
    bAsc2Sapp(&HC7) = &H4 ' "Ç"
    bAsc2Sapp(&HC8) = &H5 ' "È"
    bAsc2Sapp(&HC9) = &H6 ' "É"
    bAsc2Sapp(&HCA) = &H7 ' "Ê"
    bAsc2Sapp(&HCB) = &H8 ' "Ë"
    bAsc2Sapp(&HCC) = &H9 ' "Ì"
    bAsc2Sapp(&HCE) = &HB ' "Î"
    bAsc2Sapp(&HCF) = &HC ' "Ï"
    bAsc2Sapp(&HD2) = &HD ' "Ò"
    bAsc2Sapp(&HD3) = &HE ' "Ó"
    bAsc2Sapp(&HD4) = &HF ' "Ô"
    bAsc2Sapp(&H8C) = &H10 ' "Œ"
    bAsc2Sapp(&HD9) = &H11 ' "Ù"
    bAsc2Sapp(&HDA) = &H12 ' "Ú"
    bAsc2Sapp(&HDB) = &H13 ' "Û"
    bAsc2Sapp(&HD1) = &H14 ' "Ñ"
    bAsc2Sapp(&HDF) = &H15 ' "ß"
    bAsc2Sapp(&HE0) = &H16 ' "à"
    bAsc2Sapp(&HE1) = &H17 ' "á"
    bAsc2Sapp(&HE7) = &H19 ' "ç"
    bAsc2Sapp(&HE8) = &H1A ' "è"
    bAsc2Sapp(&HE9) = &H1B ' "é"
    bAsc2Sapp(&HEA) = &H1C ' "ê"
    bAsc2Sapp(&HEB) = &H1D ' "ë"
    bAsc2Sapp(&HEC) = &H1E ' "ì"
    bAsc2Sapp(&HEE) = &H20 ' "î"
    bAsc2Sapp(&HEF) = &H21 ' "ï"
    bAsc2Sapp(&HF2) = &H22 ' "ò"
    bAsc2Sapp(&HF3) = &H23 ' "ó"
    bAsc2Sapp(&HF4) = &H24 ' "ô"
    bAsc2Sapp(&H9C) = &H25 ' "œ"
    bAsc2Sapp(&HF9) = &H26 ' "ù"
    bAsc2Sapp(&HFA) = &H27 ' "ú"
    bAsc2Sapp(&HFB) = &H28 ' "û"
    bAsc2Sapp(&HF1) = &H29 ' "ñ"
    bAsc2Sapp(&HBA) = &H2A ' "º"
    bAsc2Sapp(&HAA) = &H2B ' "ª"
    bAsc2Sapp(&H26) = &H2D ' "&"
    bAsc2Sapp(&H2B) = &H2E ' "+"
    bAsc2Sapp(&H3D) = &H35 ' "="
    bAsc2Sapp(&H3B) = &H36 ' ";"
    bAsc2Sapp(&HBF) = &H51 ' "¿"
    bAsc2Sapp(&HA1) = &H52 ' "¡"
    bAsc2Sapp(&HCD) = &H5A ' "Í"
    bAsc2Sapp(&H25) = &H5B ' "%"
    bAsc2Sapp(&H28) = &H5C ' "("
    bAsc2Sapp(&H29) = &H5D ' ")"
    bAsc2Sapp(&HE2) = &H68 ' "â"
    bAsc2Sapp(&HED) = &H6F ' "í"
    bAsc2Sapp(&H3C) = &H85 ' "<"
    bAsc2Sapp(&H3E) = &H86 ' ">"
    bAsc2Sapp(&H30) = &HA1 ' "0"
    bAsc2Sapp(&H31) = &HA2 ' "1"
    bAsc2Sapp(&H32) = &HA3 ' "2"
    bAsc2Sapp(&H33) = &HA4 ' "3"
    bAsc2Sapp(&H34) = &HA5 ' "4"
    bAsc2Sapp(&H35) = &HA6 ' "5"
    bAsc2Sapp(&H36) = &HA7 ' "6"
    bAsc2Sapp(&H37) = &HA8 ' "7"
    bAsc2Sapp(&H38) = &HA9 ' "8"
    bAsc2Sapp(&H39) = &HAA ' "9"
    bAsc2Sapp(&H21) = &HAB ' "!"
    bAsc2Sapp(&H3F) = &HAC ' "?"
    bAsc2Sapp(&H2E) = &HAD ' "."
    bAsc2Sapp(&H2D) = &HAE ' "-"
    bAsc2Sapp(&HB7) = &HAF ' "·"
    bAsc2Sapp(&H22) = &HB2 ' """
    bAsc2Sapp(&H27) = &HB4 ' "'"
    bAsc2Sapp(&H2C) = &HB8 ' ","
    bAsc2Sapp(&H2F) = &HBA ' "/"
    bAsc2Sapp(&H41) = &HBB ' "A"
    bAsc2Sapp(&H42) = &HBC ' "B"
    bAsc2Sapp(&H43) = &HBD ' "C"
    bAsc2Sapp(&H44) = &HBE ' "D"
    bAsc2Sapp(&H45) = &HBF ' "E"
    bAsc2Sapp(&H46) = &HC0 ' "F"
    bAsc2Sapp(&H47) = &HC1 ' "G"
    bAsc2Sapp(&H48) = &HC2 ' "H"
    bAsc2Sapp(&H49) = &HC3 ' "I"
    bAsc2Sapp(&H4A) = &HC4 ' "J"
    bAsc2Sapp(&H4B) = &HC5 ' "K"
    bAsc2Sapp(&H4C) = &HC6 ' "L"
    bAsc2Sapp(&H4D) = &HC7 ' "M"
    bAsc2Sapp(&H4E) = &HC8 ' "N"
    bAsc2Sapp(&H4F) = &HC9 ' "O"
    bAsc2Sapp(&H50) = &HCA ' "P"
    bAsc2Sapp(&H51) = &HCB ' "Q"
    bAsc2Sapp(&H52) = &HCC ' "R"
    bAsc2Sapp(&H53) = &HCD ' "S"
    bAsc2Sapp(&H54) = &HCE ' "T"
    bAsc2Sapp(&H55) = &HCF ' "U"
    bAsc2Sapp(&H56) = &HD0 ' "V"
    bAsc2Sapp(&H57) = &HD1 ' "W"
    bAsc2Sapp(&H58) = &HD2 ' "X"
    bAsc2Sapp(&H59) = &HD3 ' "Y"
    bAsc2Sapp(&H5A) = &HD4 ' "Z"
    bAsc2Sapp(&H61) = &HD5 ' "a"
    bAsc2Sapp(&H62) = &HD6 ' "b"
    bAsc2Sapp(&H63) = &HD7 ' "c"
    bAsc2Sapp(&H64) = &HD8 ' "d"
    bAsc2Sapp(&H65) = &HD9 ' "e"
    bAsc2Sapp(&H66) = &HDA ' "f"
    bAsc2Sapp(&H67) = &HDB ' "g"
    bAsc2Sapp(&H68) = &HDC ' "h"
    bAsc2Sapp(&H69) = &HDD ' "i"
    bAsc2Sapp(&H6A) = &HDE ' "j"
    bAsc2Sapp(&H6B) = &HDF ' "k"
    bAsc2Sapp(&H6C) = &HE0 ' "l"
    bAsc2Sapp(&H6D) = &HE1 ' "m"
    bAsc2Sapp(&H6E) = &HE2 ' "n"
    bAsc2Sapp(&H6F) = &HE3 ' "o"
    bAsc2Sapp(&H70) = &HE4 ' "p"
    bAsc2Sapp(&H71) = &HE5 ' "q"
    bAsc2Sapp(&H72) = &HE6 ' "r"
    bAsc2Sapp(&H73) = &HE7 ' "s"
    bAsc2Sapp(&H74) = &HE8 ' "t"
    bAsc2Sapp(&H75) = &HE9 ' "u"
    bAsc2Sapp(&H76) = &HEA ' "v"
    bAsc2Sapp(&H77) = &HEB ' "w"
    bAsc2Sapp(&H78) = &HEC ' "x"
    bAsc2Sapp(&H79) = &HED ' "y"
    bAsc2Sapp(&H7A) = &HEE ' "z"
    bAsc2Sapp(&H3A) = &HF0 ' ":"
    bAsc2Sapp(&HC4) = &HF1 ' "Ä"
    bAsc2Sapp(&HD6) = &HF2 ' "Ö"
    bAsc2Sapp(&HDC) = &HF3 ' "Ü"
    bAsc2Sapp(&HE4) = &HF4 ' "ä"
    bAsc2Sapp(&HF6) = &HF5 ' "ö"
    bAsc2Sapp(&HFC) = &HF6 ' "ü"
    
    bAsc2Braille(&H20) = &H0 ' " "
    bAsc2Braille(&H41) = &H1 ' "A"
    bAsc2Braille(&H43) = &H3 ' "C"
    bAsc2Braille(&H2C) = &H4 ' ","
    bAsc2Braille(&H42) = &H5 ' "B"
    bAsc2Braille(&H49) = &H6 ' "I"
    bAsc2Braille(&H46) = &H7 ' "F"
    bAsc2Braille(&H45) = &H9 ' "E"
    bAsc2Braille(&H44) = &HB ' "D"
    bAsc2Braille(&H48) = &HD ' "H"
    bAsc2Braille(&H4A) = &HE ' "J"
    bAsc2Braille(&H47) = &HF ' "G"
    bAsc2Braille(&H4B) = &H11 ' "K"
    bAsc2Braille(&H4D) = &H13 ' "M"
    bAsc2Braille(&H4C) = &H15 ' "L"
    bAsc2Braille(&H53) = &H16 ' "S"
    bAsc2Braille(&H50) = &H17 ' "P"
    bAsc2Braille(&H4F) = &H19 ' "O"
    bAsc2Braille(&H4E) = &H1B ' "N"
    bAsc2Braille(&H52) = &H1D ' "R"
    bAsc2Braille(&H54) = &H1E ' "T"
    bAsc2Braille(&H51) = &H1F ' "Q"
    bAsc2Braille(&H2E) = &H2C ' "."
    bAsc2Braille(&H57) = &H2E ' "W"
    bAsc2Braille(&H55) = &H31 ' "U"
    bAsc2Braille(&H58) = &H33 ' "X"
    bAsc2Braille(&H56) = &H35 ' "V"
    bAsc2Braille(&H5A) = &H39 ' "Z"
    bAsc2Braille(&H59) = &H3B ' "Y"
            
    AsciiReady = True
    
End Sub

Private Sub SappInitialize()
Dim i As Long
    
    If bEnd(0) <> &HFF Then
        bEnd(0) = &HFF
    End If
    
    If cString Is Nothing Then
        Set cString = New cStringBuilder
    End If
    
    ReDim sSapp2Asc(&HFF) As String
    ReDim sSapp2AscJap(&HFF) As String
    
    For i = LBound(sSapp2Asc) To UBound(sSapp2Asc)
        sSapp2Asc(i) = "\h" & PadHex$(i, 2)
        sSapp2AscJap(i) = "\h" & PadHex$(i, 2)
    Next i
    
    sSapp2Asc(&H0) = " "
    sSapp2Asc(&H1) = "À"
    sSapp2Asc(&H2) = "Á"
    sSapp2Asc(&H3) = "Â"
    sSapp2Asc(&H4) = "Ç"
    sSapp2Asc(&H5) = "È"
    sSapp2Asc(&H6) = "É"
    sSapp2Asc(&H7) = "Ê"
    sSapp2Asc(&H8) = "Ë"
    sSapp2Asc(&H9) = "Ì"
    sSapp2Asc(&HB) = "Î"
    sSapp2Asc(&HC) = "Ï"
    sSapp2Asc(&HD) = "Ò"
    sSapp2Asc(&HE) = "Ó"
    sSapp2Asc(&HF) = "Ô"
    sSapp2Asc(&H10) = "Œ"
    sSapp2Asc(&H11) = "Ù"
    sSapp2Asc(&H12) = "Ú"
    sSapp2Asc(&H13) = "Û"
    sSapp2Asc(&H14) = "Ñ"
    sSapp2Asc(&H15) = "ß"
    sSapp2Asc(&H16) = "à"
    sSapp2Asc(&H17) = "á"
    sSapp2Asc(&H19) = "ç"
    sSapp2Asc(&H1A) = "è"
    sSapp2Asc(&H1B) = "é"
    sSapp2Asc(&H1C) = "ê"
    sSapp2Asc(&H1D) = "ë"
    sSapp2Asc(&H1E) = "ì"
    sSapp2Asc(&H20) = "î"
    sSapp2Asc(&H21) = "ï"
    sSapp2Asc(&H22) = "ò"
    sSapp2Asc(&H23) = "ó"
    sSapp2Asc(&H24) = "ô"
    sSapp2Asc(&H25) = "œ"
    sSapp2Asc(&H26) = "ù"
    sSapp2Asc(&H27) = "ú"
    sSapp2Asc(&H28) = "û"
    sSapp2Asc(&H29) = "ñ"
    sSapp2Asc(&H2A) = "º"
    sSapp2Asc(&H2B) = "ª"
    sSapp2Asc(&H2D) = "&"
    sSapp2Asc(&H2E) = "+"
    sSapp2Asc(&H34) = "[Lv]"
    sSapp2Asc(&H35) = "="
    sSapp2Asc(&H36) = ";"
    sSapp2Asc(&H51) = "¿"
    sSapp2Asc(&H52) = "¡"
    sSapp2Asc(&H53) = "[PK]"
    sSapp2Asc(&H54) = "[MN]"
    sSapp2Asc(&H55) = "[PO]"
    sSapp2Asc(&H56) = "[Ke]"
    sSapp2Asc(&H57) = "[BL]"
    sSapp2Asc(&H58) = "[OC]"
    sSapp2Asc(&H59) = "[K]"
    sSapp2Asc(&H5A) = "Í"
    sSapp2Asc(&H5B) = "%"
    sSapp2Asc(&H5C) = "("
    sSapp2Asc(&H5D) = ")"
    sSapp2Asc(&H5E) = "[Po]"
    sSapp2Asc(&H5F) = "[Ké]"
    sSapp2Asc(&H60) = "[ME]"
    sSapp2Asc(&H61) = "[LL]"
    sSapp2Asc(&H62) = "[A]"
    sSapp2Asc(&H63) = "[E]"
    sSapp2Asc(&H68) = "â"
    sSapp2Asc(&H6F) = "í"
    sSapp2Asc(&H79) = "[U]"
    sSapp2Asc(&H7A) = "[D]"
    sSapp2Asc(&H7B) = "[L]"
    sSapp2Asc(&H7C) = "[R]"
    sSapp2Asc(&H85) = "<"
    sSapp2Asc(&H86) = ">"
    sSapp2Asc(&HA1) = "0"
    sSapp2Asc(&HA2) = "1"
    sSapp2Asc(&HA3) = "2"
    sSapp2Asc(&HA4) = "3"
    sSapp2Asc(&HA5) = "4"
    sSapp2Asc(&HA6) = "5"
    sSapp2Asc(&HA7) = "6"
    sSapp2Asc(&HA8) = "7"
    sSapp2Asc(&HA9) = "8"
    sSapp2Asc(&HAA) = "9"
    sSapp2Asc(&HAB) = "!"
    sSapp2Asc(&HAC) = "?"
    sSapp2Asc(&HAD) = "."
    sSapp2Asc(&HAE) = "-"
    sSapp2Asc(&HAF) = "·"
    sSapp2Asc(&HB0) = "[.]"
    sSapp2Asc(&HB1) = "[" & ChrW$(34) & "]"
    sSapp2Asc(&HB2) = ChrW$(34)
    sSapp2Asc(&HB3) = "[']"
    sSapp2Asc(&HB4) = "'"
    sSapp2Asc(&HB5) = "[m]"
    sSapp2Asc(&HB6) = "[f]"
    sSapp2Asc(&HB7) = "[$]"
    sSapp2Asc(&HB8) = ","
    sSapp2Asc(&HB9) = "[x]"
    sSapp2Asc(&HBA) = "/"
    sSapp2Asc(&HBB) = "A"
    sSapp2Asc(&HBC) = "B"
    sSapp2Asc(&HBD) = "C"
    sSapp2Asc(&HBE) = "D"
    sSapp2Asc(&HBF) = "E"
    sSapp2Asc(&HC0) = "F"
    sSapp2Asc(&HC1) = "G"
    sSapp2Asc(&HC2) = "H"
    sSapp2Asc(&HC3) = "I"
    sSapp2Asc(&HC4) = "J"
    sSapp2Asc(&HC5) = "K"
    sSapp2Asc(&HC6) = "L"
    sSapp2Asc(&HC7) = "M"
    sSapp2Asc(&HC8) = "N"
    sSapp2Asc(&HC9) = "O"
    sSapp2Asc(&HCA) = "P"
    sSapp2Asc(&HCB) = "Q"
    sSapp2Asc(&HCC) = "R"
    sSapp2Asc(&HCD) = "S"
    sSapp2Asc(&HCE) = "T"
    sSapp2Asc(&HCF) = "U"
    sSapp2Asc(&HD0) = "V"
    sSapp2Asc(&HD1) = "W"
    sSapp2Asc(&HD2) = "X"
    sSapp2Asc(&HD3) = "Y"
    sSapp2Asc(&HD4) = "Z"
    sSapp2Asc(&HD5) = "a"
    sSapp2Asc(&HD6) = "b"
    sSapp2Asc(&HD7) = "c"
    sSapp2Asc(&HD8) = "d"
    sSapp2Asc(&HD9) = "e"
    sSapp2Asc(&HDA) = "f"
    sSapp2Asc(&HDB) = "g"
    sSapp2Asc(&HDC) = "h"
    sSapp2Asc(&HDD) = "i"
    sSapp2Asc(&HDE) = "j"
    sSapp2Asc(&HDF) = "k"
    sSapp2Asc(&HE0) = "l"
    sSapp2Asc(&HE1) = "m"
    sSapp2Asc(&HE2) = "n"
    sSapp2Asc(&HE3) = "o"
    sSapp2Asc(&HE4) = "p"
    sSapp2Asc(&HE5) = "q"
    sSapp2Asc(&HE6) = "r"
    sSapp2Asc(&HE7) = "s"
    sSapp2Asc(&HE8) = "t"
    sSapp2Asc(&HE9) = "u"
    sSapp2Asc(&HEA) = "v"
    sSapp2Asc(&HEB) = "w"
    sSapp2Asc(&HEC) = "x"
    sSapp2Asc(&HED) = "y"
    sSapp2Asc(&HEE) = "z"
    sSapp2Asc(&HEF) = "[>]"
    sSapp2Asc(&HF0) = ":"
    sSapp2Asc(&HF1) = "Ä"
    sSapp2Asc(&HF2) = "Ö"
    sSapp2Asc(&HF3) = "Ü"
    sSapp2Asc(&HF4) = "ä"
    sSapp2Asc(&HF5) = "ö"
    sSapp2Asc(&HF6) = "ü"
    'sSapp2Asc(&HF7) = "[u]"
    'sSapp2Asc(&HF8) = "[d]"
    'sSapp2Asc(&HF9) = "[l]"
    sSapp2Asc(&HFA) = "\l"
    sSapp2Asc(&HFB) = "\p"
    sSapp2Asc(&HFC) = "\c"
    sSapp2Asc(&HFD) = "\v"
    sSapp2Asc(&HFE) = "\n"
    sSapp2Asc(&HFF) = vbNullString ' "\x"
    
    sSapp2AscJap(&H0) = sSapp2Asc(&H0)
    sSapp2AscJap(&H1) = "a"
    sSapp2AscJap(&H2) = "i"
    sSapp2AscJap(&H3) = "u"
    sSapp2AscJap(&H4) = "e"
    sSapp2AscJap(&H5) = "o"
    sSapp2AscJap(&H6) = "ka"
    sSapp2AscJap(&H7) = "ki"
    sSapp2AscJap(&H8) = "ku"
    sSapp2AscJap(&H9) = "ke"
    sSapp2AscJap(&HA) = "ko"
    sSapp2AscJap(&HB) = "sa"
    sSapp2AscJap(&HC) = "shi"
    sSapp2AscJap(&HD) = "su"
    sSapp2AscJap(&HE) = "se"
    sSapp2AscJap(&HF) = "so"
    sSapp2AscJap(&H10) = "ta"
    sSapp2AscJap(&H11) = "chi"
    sSapp2AscJap(&H12) = "tsu"
    sSapp2AscJap(&H13) = "te"
    sSapp2AscJap(&H14) = "to"
    sSapp2AscJap(&H15) = "na"
    sSapp2AscJap(&H16) = "ni"
    sSapp2AscJap(&H17) = "nu"
    sSapp2AscJap(&H18) = "ne"
    sSapp2AscJap(&H19) = "no"
    sSapp2AscJap(&H1A) = "ha"
    sSapp2AscJap(&H1B) = "hi"
    sSapp2AscJap(&H1C) = "fu"
    sSapp2AscJap(&H1D) = "he"
    sSapp2AscJap(&H1E) = "ho"
    sSapp2AscJap(&H1F) = "ma"
    sSapp2AscJap(&H20) = "mi"
    sSapp2AscJap(&H21) = "mu"
    sSapp2AscJap(&H22) = "me"
    sSapp2AscJap(&H23) = "mo"
    sSapp2AscJap(&H24) = "ya"
    sSapp2AscJap(&H25) = "yu"
    sSapp2AscJap(&H26) = "yo"
    sSapp2AscJap(&H27) = "ra"
    sSapp2AscJap(&H28) = "ri"
    sSapp2AscJap(&H29) = "ru"
    sSapp2AscJap(&H2A) = "re"
    sSapp2AscJap(&H2B) = "ro"
    sSapp2AscJap(&H2C) = "wa"
    sSapp2AscJap(&H2D) = "wo"
    sSapp2AscJap(&H2E) = "n"
    sSapp2AscJap(&H2F) = "a"
    sSapp2AscJap(&H30) = "i"
    sSapp2AscJap(&H31) = "u"
    sSapp2AscJap(&H32) = "e"
    sSapp2AscJap(&H33) = "o"
    sSapp2AscJap(&H34) = "ya"
    sSapp2AscJap(&H35) = "yu"
    sSapp2AscJap(&H36) = "yo"
    sSapp2AscJap(&H37) = "ga"
    sSapp2AscJap(&H38) = "gi"
    sSapp2AscJap(&H39) = "gu"
    sSapp2AscJap(&H3A) = "ge"
    sSapp2AscJap(&H3B) = "go"
    sSapp2AscJap(&H3C) = "za"
    sSapp2AscJap(&H3D) = "ji"
    sSapp2AscJap(&H3E) = "zu"
    sSapp2AscJap(&H3F) = "ze"
    sSapp2AscJap(&H40) = "zo"
    sSapp2AscJap(&H41) = "da"
    sSapp2AscJap(&H42) = "dji"
    sSapp2AscJap(&H43) = "du"
    sSapp2AscJap(&H44) = "de"
    sSapp2AscJap(&H45) = "do"
    sSapp2AscJap(&H46) = "ba"
    sSapp2AscJap(&H47) = "bi"
    sSapp2AscJap(&H48) = "bu"
    sSapp2AscJap(&H49) = "be"
    sSapp2AscJap(&H4A) = "bo"
    sSapp2AscJap(&H4B) = "pa"
    sSapp2AscJap(&H4C) = "pi"
    sSapp2AscJap(&H4D) = "pu"
    sSapp2AscJap(&H4E) = "pe"
    sSapp2AscJap(&H4F) = "po"
    sSapp2AscJap(&H50) = "tsu"
    sSapp2AscJap(&H51) = "A"
    sSapp2AscJap(&H52) = "I"
    sSapp2AscJap(&H53) = "U"
    sSapp2AscJap(&H54) = "E"
    sSapp2AscJap(&H55) = "O"
    sSapp2AscJap(&H56) = "KA"
    sSapp2AscJap(&H57) = "KI"
    sSapp2AscJap(&H58) = "KU"
    sSapp2AscJap(&H59) = "KE"
    sSapp2AscJap(&H5A) = "KO"
    sSapp2AscJap(&H5B) = "SA"
    sSapp2AscJap(&H5C) = "SHI"
    sSapp2AscJap(&H5D) = "SU"
    sSapp2AscJap(&H5E) = "SE"
    sSapp2AscJap(&H5F) = "SO"
    sSapp2AscJap(&H60) = "TA"
    sSapp2AscJap(&H61) = "CHI"
    sSapp2AscJap(&H62) = "TSU"
    sSapp2AscJap(&H63) = "TE"
    sSapp2AscJap(&H64) = "TO"
    sSapp2AscJap(&H65) = "NA"
    sSapp2AscJap(&H66) = "NI"
    sSapp2AscJap(&H67) = "NU"
    sSapp2AscJap(&H68) = "NE"
    sSapp2AscJap(&H69) = "NO"
    sSapp2AscJap(&H6A) = "HA"
    sSapp2AscJap(&H6B) = "HI"
    sSapp2AscJap(&H6C) = "FU"
    sSapp2AscJap(&H6D) = "HE"
    sSapp2AscJap(&H6E) = "HO"
    sSapp2AscJap(&H6F) = "MA"
    sSapp2AscJap(&H70) = "MI"
    sSapp2AscJap(&H71) = "MU"
    sSapp2AscJap(&H72) = "ME"
    sSapp2AscJap(&H73) = "MO"
    sSapp2AscJap(&H74) = "YA"
    sSapp2AscJap(&H75) = "YU"
    sSapp2AscJap(&H76) = "YO"
    sSapp2AscJap(&H77) = "RA"
    sSapp2AscJap(&H78) = "RI"
    sSapp2AscJap(&H79) = "RU"
    sSapp2AscJap(&H7A) = "RE"
    sSapp2AscJap(&H7B) = "RO"
    sSapp2AscJap(&H7C) = "WA"
    sSapp2AscJap(&H7D) = "WO"
    sSapp2AscJap(&H7E) = "N"
    sSapp2AscJap(&H7F) = "A"
    sSapp2AscJap(&H80) = "I"
    sSapp2AscJap(&H81) = "U"
    sSapp2AscJap(&H82) = "E"
    sSapp2AscJap(&H83) = "O"
    sSapp2AscJap(&H84) = "YA"
    sSapp2AscJap(&H85) = "YU"
    sSapp2AscJap(&H86) = "YO"
    sSapp2AscJap(&H87) = "GA"
    sSapp2AscJap(&H88) = "GI"
    sSapp2AscJap(&H89) = "GU"
    sSapp2AscJap(&H8A) = "GE"
    sSapp2AscJap(&H8B) = "GO"
    sSapp2AscJap(&H8C) = "ZA"
    sSapp2AscJap(&H8D) = "JI"
    sSapp2AscJap(&H8E) = "ZU"
    sSapp2AscJap(&H8F) = "ZE"
    sSapp2AscJap(&H90) = "ZO"
    sSapp2AscJap(&H91) = "DA"
    sSapp2AscJap(&H92) = "DJI"
    sSapp2AscJap(&H93) = "DU"
    sSapp2AscJap(&H94) = "DE"
    sSapp2AscJap(&H95) = "DO"
    sSapp2AscJap(&H96) = "BA"
    sSapp2AscJap(&H97) = "BI"
    sSapp2AscJap(&H98) = "BU"
    sSapp2AscJap(&H99) = "BE"
    sSapp2AscJap(&H9A) = "BO"
    sSapp2AscJap(&H9B) = "PA"
    sSapp2AscJap(&H9C) = "PI"
    sSapp2AscJap(&H9D) = "PU"
    sSapp2AscJap(&H9E) = "PE"
    sSapp2AscJap(&H9F) = "PO"
    sSapp2AscJap(&HA0) = "TSU"
    sSapp2AscJap(&HA1) = sSapp2Asc(&HA1) ' "0"
    sSapp2AscJap(&HA2) = sSapp2Asc(&HA2) ' "1"
    sSapp2AscJap(&HA3) = sSapp2Asc(&HA3) ' "2"
    sSapp2AscJap(&HA4) = sSapp2Asc(&HA4) ' "3"
    sSapp2AscJap(&HA5) = sSapp2Asc(&HA5) ' "4"
    sSapp2AscJap(&HA6) = sSapp2Asc(&HA6) ' "5"
    sSapp2AscJap(&HA7) = sSapp2Asc(&HA7) ' "6"
    sSapp2AscJap(&HA8) = sSapp2Asc(&HA8) ' "7"
    sSapp2AscJap(&HA9) = sSapp2Asc(&HA9) ' "8"
    sSapp2AscJap(&HAA) = sSapp2Asc(&HAA) ' "9"
    sSapp2AscJap(&HAB) = sSapp2Asc(&HAB) ' "!"
    sSapp2AscJap(&HAC) = sSapp2Asc(&HAC) ' "?"
    sSapp2AscJap(&HAD) = sSapp2Asc(&HAD) ' "."
    sSapp2AscJap(&HAE) = sSapp2Asc(&HAE) ' "-"
    sSapp2AscJap(&HAF) = sSapp2Asc(&HAF) ' "·"
    sSapp2AscJap(&HB0) = sSapp2Asc(&HB0) ' "[.]"
    sSapp2AscJap(&HB1) = sSapp2Asc(&HB1) ' "["]"
    sSapp2AscJap(&HB2) = sSapp2Asc(&HB2) ' """
    sSapp2AscJap(&HB3) = sSapp2Asc(&HB3) ' "[']"
    sSapp2AscJap(&HB4) = sSapp2Asc(&HB4) ' "'"
    sSapp2AscJap(&HB5) = sSapp2Asc(&HB5) ' "[m]"
    sSapp2AscJap(&HB6) = sSapp2Asc(&HB6) ' "[f]"
    sSapp2AscJap(&HB7) = sSapp2Asc(&HB7) ' "[$]"
    sSapp2AscJap(&HB8) = sSapp2Asc(&HB8) ' ","
    sSapp2AscJap(&HB9) = sSapp2Asc(&HB9) ' "[x]"
    sSapp2AscJap(&HBA) = sSapp2Asc(&HBA) ' "/"
    sSapp2AscJap(&HBB) = sSapp2Asc(&HBB) ' "A"
    sSapp2AscJap(&HBC) = sSapp2Asc(&HBC) ' "B"
    sSapp2AscJap(&HBD) = sSapp2Asc(&HBD) ' "C"
    sSapp2AscJap(&HBE) = sSapp2Asc(&HBE) ' "D"
    sSapp2AscJap(&HBF) = sSapp2Asc(&HBF) ' "E"
    sSapp2AscJap(&HC0) = sSapp2Asc(&HC0) ' "F"
    sSapp2AscJap(&HC1) = sSapp2Asc(&HC1) ' "G"
    sSapp2AscJap(&HC2) = sSapp2Asc(&HC2) ' "H"
    sSapp2AscJap(&HC3) = sSapp2Asc(&HC3) ' "I"
    sSapp2AscJap(&HC4) = sSapp2Asc(&HC4) ' "J"
    sSapp2AscJap(&HC5) = sSapp2Asc(&HC5) ' "K"
    sSapp2AscJap(&HC6) = sSapp2Asc(&HC6) ' "L"
    sSapp2AscJap(&HC7) = sSapp2Asc(&HC7) ' "M"
    sSapp2AscJap(&HC8) = sSapp2Asc(&HC8) ' "N"
    sSapp2AscJap(&HC9) = sSapp2Asc(&HC9) ' "O"
    sSapp2AscJap(&HCA) = sSapp2Asc(&HCA) ' "P"
    sSapp2AscJap(&HCB) = sSapp2Asc(&HCB) ' "Q"
    sSapp2AscJap(&HCC) = sSapp2Asc(&HCC) ' "R"
    sSapp2AscJap(&HCD) = sSapp2Asc(&HCD) ' "S"
    sSapp2AscJap(&HCE) = sSapp2Asc(&HCE) ' "T"
    sSapp2AscJap(&HCF) = sSapp2Asc(&HCF) ' "U"
    sSapp2AscJap(&HD0) = sSapp2Asc(&HD0) ' "V"
    sSapp2AscJap(&HD1) = sSapp2Asc(&HD1) ' "W"
    sSapp2AscJap(&HD2) = sSapp2Asc(&HD2) ' "X"
    sSapp2AscJap(&HD3) = sSapp2Asc(&HD3) ' "Y"
    sSapp2AscJap(&HD4) = sSapp2Asc(&HD4) ' "Z"
    sSapp2AscJap(&HD5) = sSapp2Asc(&HD5) ' "a"
    sSapp2AscJap(&HD6) = sSapp2Asc(&HD6) ' "b"
    sSapp2AscJap(&HD7) = sSapp2Asc(&HD7) ' "c"
    sSapp2AscJap(&HD8) = sSapp2Asc(&HD8) ' "d"
    sSapp2AscJap(&HD9) = sSapp2Asc(&HD9) ' "e"
    sSapp2AscJap(&HDA) = sSapp2Asc(&HDA) ' "f"
    sSapp2AscJap(&HDB) = sSapp2Asc(&HDB) ' "g"
    sSapp2AscJap(&HDC) = sSapp2Asc(&HDC) ' "h"
    sSapp2AscJap(&HDD) = sSapp2Asc(&HDD) ' "i"
    sSapp2AscJap(&HDE) = sSapp2Asc(&HDE) ' "j"
    sSapp2AscJap(&HDF) = sSapp2Asc(&HDF) ' "k"
    sSapp2AscJap(&HE0) = sSapp2Asc(&HE0) ' "l"
    sSapp2AscJap(&HE1) = sSapp2Asc(&HE1) ' "m"
    sSapp2AscJap(&HE2) = sSapp2Asc(&HE2) ' "n"
    sSapp2AscJap(&HE3) = sSapp2Asc(&HE3) ' "o"
    sSapp2AscJap(&HE4) = sSapp2Asc(&HE4) ' "p"
    sSapp2AscJap(&HE5) = sSapp2Asc(&HE5) ' "q"
    sSapp2AscJap(&HE6) = sSapp2Asc(&HE6) ' "r"
    sSapp2AscJap(&HE7) = sSapp2Asc(&HE7) ' "s"
    sSapp2AscJap(&HE8) = sSapp2Asc(&HE8) ' "t"
    sSapp2AscJap(&HE9) = sSapp2Asc(&HE9) ' "u"
    sSapp2AscJap(&HEA) = sSapp2Asc(&HEA) ' "v"
    sSapp2AscJap(&HEB) = sSapp2Asc(&HEB) ' "w"
    sSapp2AscJap(&HEC) = sSapp2Asc(&HEC) ' "x"
    sSapp2AscJap(&HED) = sSapp2Asc(&HED) ' "y"
    sSapp2AscJap(&HEE) = sSapp2Asc(&HEE) ' "z"
    sSapp2AscJap(&HEF) = sSapp2Asc(&HEF) ' "[>]"
    sSapp2AscJap(&HF0) = sSapp2Asc(&HF0) ' ":"
    sSapp2AscJap(&HF1) = sSapp2Asc(&HF1) ' "Ä"
    sSapp2AscJap(&HF2) = sSapp2Asc(&HF2) ' "Ö"
    sSapp2AscJap(&HF3) = sSapp2Asc(&HF3) ' "Ü"
    sSapp2AscJap(&HF4) = sSapp2Asc(&HF4) ' "ä"
    sSapp2AscJap(&HF5) = sSapp2Asc(&HF5) ' "ö"
    sSapp2AscJap(&HF6) = sSapp2Asc(&HF6) ' "ü"
    'sSapp2AscJap(&HF7) = sSapp2Asc(&HF7) ' "[u]"
    'sSapp2AscJap(&HF8) = sSapp2Asc(&HF8) ' "[d]"
    'sSapp2AscJap(&HF9) = sSapp2Asc(&HF9) ' "[l]"
    sSapp2AscJap(&HFA) = sSapp2Asc(&HFA) ' "\l"
    sSapp2AscJap(&HFB) = sSapp2Asc(&HFB) ' "\p"
    sSapp2AscJap(&HFC) = sSapp2Asc(&HFC) ' "\c"
    sSapp2AscJap(&HFD) = sSapp2Asc(&HFD) ' "\v"
    sSapp2AscJap(&HFE) = sSapp2Asc(&HFE) ' "\n"
    sSapp2AscJap(&HFF) = sSapp2Asc(&HFF) ' "\x"
    
    SappReady = True

End Sub

Private Sub BrailleInitialize()
Dim i As Long

    If bEnd(0) <> &HFF Then
        bEnd(0) = &HFF
    End If
    
    If cString Is Nothing Then
        Set cString = New cStringBuilder
    End If

    ReDim sBraille2Asc(&HFF) As String
    
    For i = LBound(sBraille2Asc) To UBound(sBraille2Asc)
        sBraille2Asc(i) = "\h" & PadHex$(i, 2)
    Next i
    
    sBraille2Asc(&H0) = " "
    sBraille2Asc(&H1) = "A"
    sBraille2Asc(&H3) = "C"
    sBraille2Asc(&H4) = ","
    sBraille2Asc(&H5) = "B"
    sBraille2Asc(&H6) = "I"
    sBraille2Asc(&H7) = "F"
    sBraille2Asc(&H9) = "E"
    sBraille2Asc(&HB) = "D"
    sBraille2Asc(&HD) = "H"
    sBraille2Asc(&HE) = "J"
    sBraille2Asc(&HF) = "G"
    sBraille2Asc(&H11) = "K"
    sBraille2Asc(&H13) = "M"
    sBraille2Asc(&H15) = "L"
    sBraille2Asc(&H16) = "S"
    sBraille2Asc(&H17) = "P"
    sBraille2Asc(&H19) = "O"
    sBraille2Asc(&H1B) = "N"
    sBraille2Asc(&H1D) = "R"
    sBraille2Asc(&H1E) = "T"
    sBraille2Asc(&H1F) = "Q"
    sBraille2Asc(&H2C) = "."
    sBraille2Asc(&H2E) = "W"
    sBraille2Asc(&H31) = "U"
    sBraille2Asc(&H33) = "X"
    sBraille2Asc(&H35) = "V"
    sBraille2Asc(&H39) = "Z"
    sBraille2Asc(&H3B) = "Y"
    sBraille2Asc(&HFE) = "\n"
    sBraille2Asc(&HFF) = vbNullString ' "\x"
    
    BrailleReady = True

End Sub

Public Function Asc2SappLen(ByVal sAscii As String) As Long
Dim iChars() As Integer
Dim Skip As Boolean
Dim lPointer As Long
Dim lSavePtr As Long
Dim lUBound As Long
Dim i As Long
    
    If Not AsciiReady Then
        AscInitialize
    End If
    
    If InStrB(1, sAscii, "[") Then
            
        If InStrB(sAscii, "]") - InStrB(sAscii, "[") \ 2& - 1& > 2& Then
    
            DoReplace sAscii, "[player]", "\v\h01", , , vbTextCompare
            DoReplace sAscii, "[buffer1]", "\v\h02", , , vbTextCompare
            DoReplace sAscii, "[buffer2]", "\v\h03", , , vbTextCompare
            DoReplace sAscii, "[buffer3]", "\v\h04", , , vbTextCompare
            DoReplace sAscii, "[rival]", "\v\h06", , , vbTextCompare
            DoReplace sAscii, "[game]", "\v\h07", , , vbTextCompare
            DoReplace sAscii, "[team]", "\v\h08", , , vbTextCompare
            DoReplace sAscii, "[otherteam]", "\v\h09", , , vbTextCompare
            DoReplace sAscii, "[teamleader]", "\v\h0A", , , vbTextCompare
            DoReplace sAscii, "[otherteamleader]", "\v\h0B", , , vbTextCompare
            DoReplace sAscii, "[legend]", "\v\h0C", , , vbTextCompare
            DoReplace sAscii, "[otherlegend]", "\v\h0D", , , vbTextCompare
            DoReplace sAscii, "[small]", "\c\h06\h00", , , vbTextCompare
            DoReplace sAscii, "[jap]", "\c\h15", , , vbTextCompare
            DoReplace sAscii, "[west]", "\c\h16", , , vbTextCompare
            
            If InStrB(1, sAscii, "_rs") Then
                
                DoReplace sAscii, "[transp_rs]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[darkgrey_rs]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[red_rs]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[lightgreen_rs]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[blue_rs]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[yellow_rs]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[cyan_rs]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[magenta_rs]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[grey_rs]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[black_rs]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[black2_rs]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[lightgrey_rs]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[white_rs]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[skyblue_rs]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[darkskyblue_rs]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[white2_rs]", "\c\h01\h0F", , , vbTextCompare
                
                DoReplace sAscii, "[black_rst]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[white_rst]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[red_rst]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[navy_rst]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[lightnavy_rst]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[white2_rst]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[darkpurple_rst]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[lightpurple_rst]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[darknavy_rst]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[darkgrey_rst]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[grey_rst]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[darkbronze_rst]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[bronze_rst]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[darkred_rst]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[purple_rst]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[transp_rst]", "\c\h01\h0F", , , vbTextCompare
                
            ElseIf InStrB(1, sAscii, "_fr") Then
                
                DoReplace sAscii, "[white_fr]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[white2_fr]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[black_fr]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[grey_fr]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[red_fr]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[orange_fr]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[green_fr]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[lightgreen_fr]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[blue_fr]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[lightblue_fr]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[white3_fr]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[lightblue2_fr]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[cyan_fr]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[lightblue3_fr]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[navyblue_fr]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[darknavyblue_fr]", "\c\h01\h0F", , , vbTextCompare
                
                DoReplace sAscii, "[darknavy_frt]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[white_frt]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[red_frt]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[navy_frt]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[lightnavy_frt]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[white_frt]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[purple2_frt]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[grey_frt]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[darkpurple_frt]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[black_frt]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[magenta_frt]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[white2_frt]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[magenta_frt]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[gold_frt]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[lightgold_frt]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[darknavy2_frt]", "\c\h01\h0F", , , vbTextCompare
            
            ElseIf InStrB(1, sAscii, "_em") Then
            
                DoReplace sAscii, "[white_em]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[white2_em]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[darkgrey_em]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[grey_em]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[red_em]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[orange_em]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[green_em]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[lightgreen_em]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[blue_em]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[lightblue_em]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[white3_em]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[white4_em]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[white5_em]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[limegreen_em]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[aqua_em]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[navy_em]", "\c\h01\h0F", , , vbTextCompare
                
                DoReplace sAscii, "[transp_emt]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[white_emt]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[red_emt]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[darknavy_emt]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[navy_emt]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[white2_emt]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[grey_emt]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[lightgrey_emt]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[darknavy2_emt]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[darkgrey_emt]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[lightgrey_emt]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[brown_emt]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[lightbrown_emt]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[darkred_emt]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[purple_emt]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[transp2_emt]", "\c\h01\h0F", , , vbTextCompare
            
            End If
            
        End If
        
    End If
    
    ReDim iChars(1& To 1&) As Integer
    lSavePtr = VarPtr(iChars(1))
    
    RtlMoveMemory ByVal VarPtr(lPointer), ByVal ArrPtr(iChars), 4&
    RtlMoveMemory ByVal lPointer + 16&, &H7FFFFFFF, 4&
    lPointer = lPointer + 12&
    
    RtlMoveMemory ByVal lPointer, StrPtr(sAscii), 4&
    lUBound = Len(sAscii)
    
    For i = LBound(iChars) To lUBound

        If iChars(i) <> &H5B Then ' "["
            
            If iChars(i) = &H5C Then ' "\"
                
                If i + 1& <= lUBound Then
                
                    Select Case iChars(i + 1&)
                      
                        Case &H6E, &H4E: Skip = True ' "n"
                        Case &H70, &H50: Skip = True ' "p"
                        
                        Case &H68, &H48 ' "h"
                        
                            If i + 3& <= lUBound Then
                                If IsHexI(iChars(i + 2&), iChars(i + 3&)) Then
                                    i = i + 3&
                                End If
                            End If
                        
                        Case &H76, &H56: Skip = True ' "v"
                        Case &H6C, &H4C: Skip = True ' "l"
                        Case &H63, &H43: Skip = True ' "c"
                        Case &H78, &H58: Skip = True: Exit For ' "x"
                          
                  End Select
        
                  If Skip = True Then
                      i = i + 1&
                  End If
                  
                  Skip = False
                
                End If
                
            End If
            
        Else
            
            If i + 3& <= lUBound Then
            
                If iChars(i + 3&) = &H5D Then ' "]"
                
                    Select Case iChars(i + 1&)
                        
                        Case &H4D ' "M"
                            If iChars(i + 2&) = &H45 Then ' "E"
                                Skip = True
                            ElseIf iChars(i + 2&) = &H4E Then ' "N"
                                Skip = True
                            End If
                        
                        Case &H4B ' "K"
                            If iChars(i + 2&) = &H65 Then ' "e"
                                Skip = True
                            ElseIf iChars(i + 2&) = &HE9 Then ' "é"
                                Skip = True
                            End If
                        
                        Case &H4C ' "L"
                            If iChars(i + 2&) = &H76 Then ' "v"
                                Skip = True
                            ElseIf iChars(i + 2&) = &H4C Then ' "L"
                                Skip = True
                            End If
                            
                        Case &H50 ' "P"
                            If iChars(i + 2&) = &H4B Then ' "K"
                                Skip = True
                            ElseIf iChars(i + 2&) = &H4F Then ' "O"
                                Skip = True
                            ElseIf iChars(i + 2&) = &H6F Then ' "o"
                                Skip = True
                            End If
                            
                        Case &H42 ' "B"
                            If iChars(i + 2&) = &H4C Then ' "L"
                                Skip = True
                            End If
                        
                        Case &H4F ' "O"
                            If iChars(i + 2&) = &H43 Then ' "C"
                                Skip = True
                            End If
                    
                    End Select
                    
                    If Skip = True Then
                        i = i + 3&
                    End If
                    
                End If
            
            End If
            
            If Skip = False Then
            
                If i + 2& <= lUBound Then
            
                    If iChars(i + 2&) = &H5D Then ' "]"
                
                        Select Case iChars(i + 1&)
                            Case &H2E: Skip = True ' "."
                            Case &H22: Skip = True ' """
                            Case &H24: Skip = True ' "$"
                            Case &H6D: Skip = True ' "m"
                            Case &H66: Skip = True ' "f"
                            Case &H4B: Skip = True ' "K"
                            Case &H41: Skip = True ' "A"
                            Case &H45: Skip = True ' "E"
                            Case &H55: Skip = True ' "U"
                            Case &H44: Skip = True ' "D"
                            Case &H4C: Skip = True ' "L"
                            Case &H52: Skip = True ' "R"
                            Case &H27: Skip = True ' "'"
                            Case &H78: Skip = True ' "x"
                            Case &H3E: Skip = True ' ">"
                        End Select
                
                        If Skip = True Then
                            i = i + 2&
                        End If
                    
                    End If
                
                End If
                
            End If
            
            Skip = False
            
        End If
        
        Asc2SappLen = Asc2SappLen + 1&
        
    Next i
    
    RtlMoveMemory ByVal lPointer, lSavePtr, 4&
    
End Function

Public Sub Asc2Sapp(ByVal sAscii As String, ByRef bArray() As Byte)
Dim iChars() As Integer
Dim Skip As Boolean
Dim lCounter As Long
Dim lPointer As Long
Dim lSavePtr As Long
Dim lUBound As Long
Dim i As Long

    If Not AsciiReady Then
        AscInitialize
    End If
       
    If InStrB(1, sAscii, "[") Then
            
        If InStrB(sAscii, "]") - InStrB(sAscii, "[") \ 2& - 1& > 2& Then
    
            DoReplace sAscii, "[player]", "\v\h01", , , vbTextCompare
            DoReplace sAscii, "[buffer1]", "\v\h02", , , vbTextCompare
            DoReplace sAscii, "[buffer2]", "\v\h03", , , vbTextCompare
            DoReplace sAscii, "[buffer3]", "\v\h04", , , vbTextCompare
            DoReplace sAscii, "[rival]", "\v\h06", , , vbTextCompare
            DoReplace sAscii, "[game]", "\v\h07", , , vbTextCompare
            DoReplace sAscii, "[team]", "\v\h08", , , vbTextCompare
            DoReplace sAscii, "[otherteam]", "\v\h09", , , vbTextCompare
            DoReplace sAscii, "[teamleader]", "\v\h0A", , , vbTextCompare
            DoReplace sAscii, "[otherteamleader]", "\v\h0B", , , vbTextCompare
            DoReplace sAscii, "[legend]", "\v\h0C", , , vbTextCompare
            DoReplace sAscii, "[otherlegend]", "\v\h0D", , , vbTextCompare
            DoReplace sAscii, "[small]", "\c\h06\h00", , , vbTextCompare
            DoReplace sAscii, "[jap]", "\c\h15", , , vbTextCompare
            DoReplace sAscii, "[west]", "\c\h16", , , vbTextCompare
            
            If InStrB(1, sAscii, "_rs") Then
                
                DoReplace sAscii, "[transp_rs]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[darkgrey_rs]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[red_rs]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[lightgreen_rs]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[blue_rs]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[yellow_rs]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[cyan_rs]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[magenta_rs]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[grey_rs]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[black_rs]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[black2_rs]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[lightgrey_rs]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[white_rs]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[skyblue_rs]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[darkskyblue_rs]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[white2_rs]", "\c\h01\h0F", , , vbTextCompare
                
                DoReplace sAscii, "[black_rst]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[white_rst]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[red_rst]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[navy_rst]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[lightnavy_rst]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[white2_rst]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[darkpurple_rst]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[lightpurple_rst]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[darknavy_rst]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[darkgrey_rst]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[grey_rst]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[darkbronze_rst]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[bronze_rst]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[darkred_rst]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[purple_rst]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[transp_rst]", "\c\h01\h0F", , , vbTextCompare
                
            ElseIf InStrB(1, sAscii, "_fr") Then
                
                DoReplace sAscii, "[white_fr]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[white2_fr]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[black_fr]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[grey_fr]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[red_fr]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[orange_fr]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[green_fr]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[lightgreen_fr]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[blue_fr]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[lightblue_fr]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[white3_fr]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[lightblue2_fr]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[cyan_fr]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[lightblue3_fr]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[navyblue_fr]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[darknavyblue_fr]", "\c\h01\h0F", , , vbTextCompare
                
                DoReplace sAscii, "[darknavy_frt]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[white_frt]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[red_frt]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[navy_frt]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[lightnavy_frt]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[white_frt]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[purple2_frt]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[grey_frt]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[darkpurple_frt]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[black_frt]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[magenta_frt]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[white2_frt]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[magenta_frt]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[gold_frt]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[lightgold_frt]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[darknavy2_frt]", "\c\h01\h0F", , , vbTextCompare
            
            ElseIf InStrB(1, sAscii, "_em") Then
            
                DoReplace sAscii, "[white_em]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[white2_em]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[darkgrey_em]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[grey_em]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[red_em]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[orange_em]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[green_em]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[lightgreen_em]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[blue_em]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[lightblue_em]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[white3_em]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[white4_em]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[white5_em]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[limegreen_em]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[aqua_em]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[navy_em]", "\c\h01\h0F", , , vbTextCompare
                
                DoReplace sAscii, "[transp_emt]", "\c\h01\h00", , , vbTextCompare
                DoReplace sAscii, "[white_emt]", "\c\h01\h01", , , vbTextCompare
                DoReplace sAscii, "[red_emt]", "\c\h01\h02", , , vbTextCompare
                DoReplace sAscii, "[darknavy_emt]", "\c\h01\h03", , , vbTextCompare
                DoReplace sAscii, "[navy_emt]", "\c\h01\h04", , , vbTextCompare
                DoReplace sAscii, "[white2_emt]", "\c\h01\h05", , , vbTextCompare
                DoReplace sAscii, "[grey_emt]", "\c\h01\h06", , , vbTextCompare
                DoReplace sAscii, "[lightgrey_emt]", "\c\h01\h07", , , vbTextCompare
                DoReplace sAscii, "[darknavy2_emt]", "\c\h01\h08", , , vbTextCompare
                DoReplace sAscii, "[darkgrey_emt]", "\c\h01\h09", , , vbTextCompare
                DoReplace sAscii, "[lightgrey_emt]", "\c\h01\h0A", , , vbTextCompare
                DoReplace sAscii, "[brown_emt]", "\c\h01\h0B", , , vbTextCompare
                DoReplace sAscii, "[lightbrown_emt]", "\c\h01\h0C", , , vbTextCompare
                DoReplace sAscii, "[darkred_emt]", "\c\h01\h0D", , , vbTextCompare
                DoReplace sAscii, "[purple_emt]", "\c\h01\h0E", , , vbTextCompare
                DoReplace sAscii, "[transp2_emt]", "\c\h01\h0F", , , vbTextCompare
            
            End If
            
        End If
        
    End If
    
    'bTempArray() = sAscii
    ReDim iChars(1& To 1&) As Integer
    lSavePtr = VarPtr(iChars(1))
    
    RtlMoveMemory ByVal VarPtr(lPointer), ByVal ArrPtr(iChars), 4&
    RtlMoveMemory ByVal lPointer + 16&, &H7FFFFFFF, 4&
    lPointer = lPointer + 12&
    
    RtlMoveMemory ByVal lPointer, StrPtr(sAscii), 4&
    lUBound = Len(sAscii)
    
    For i = LBound(iChars) To lUBound
    
        If iChars(i) <> &H5B Then ' "["
            
            If iChars(i) <> &H5C Then ' "\"
                bArray(lCounter) = bAsc2Sapp(iChars(i))
            Else
                
                If i + 1& <= lUBound Then
                
                    Select Case iChars(i + 1&)
                    
                        Case &H6E, &H4E: bArray(lCounter) = &HFE: Skip = True ' "n"
                        Case &H78, &H58: bArray(lCounter) = &HFF: Skip = True: Exit For ' "x"
                        Case &H70, &H50: bArray(lCounter) = &HFB: Skip = True ' "p"
                        
                        Case &H68, &H48 ' "h"
                        
                            If i + 3& <= lUBound Then
                                If IsHexI(iChars(i + 2&), iChars(i + 3&)) Then
                                    bArray(lCounter) = CByte("&H" & ChrW$(iChars(i + 2&)) & ChrW$(iChars(i + 3&)))
                                    i = i + 3&
                                End If
                            End If
                        
                        Case &H76, &H56: bArray(lCounter) = &HFD: Skip = True ' "v"
                        Case &H6C, &H4C: bArray(lCounter) = &HFA: Skip = True ' "l"
                        Case &H63, &H43: bArray(lCounter) = &HFC: Skip = True ' "c"

                    End Select
                    
                    If Skip = True Then
                        i = i + 1&
                    End If
                    
                    Skip = False
                
                End If
                
            End If
            
        Else
            
            If i + 3& <= lUBound Then
            
                If iChars(i + 3&) = &H5D Then ' "]"
                
                    Select Case iChars(i + 1&)
                        
                        Case &H4D ' "M"
                            If iChars(i + 2&) = &H45 Then ' "E"
                                bArray(lCounter) = &H60: Skip = True
                            ElseIf iChars(i + 2&) = &H4E Then ' "N"
                                bArray(lCounter) = &H54: Skip = True
                            End If
                        
                        Case &H4B ' "K"
                            If iChars(i + 2&) = &H65 Then ' "e"
                                bArray(lCounter) = &H56: Skip = True
                            ElseIf iChars(i + 2&) = &HE9 Then ' "é"
                                bArray(lCounter) = &H5F: Skip = True
                            End If
                        
                        Case &H4C ' "L"
                            If iChars(i + 2&) = &H76 Then ' "v"
                                bArray(lCounter) = &H34: Skip = True
                            ElseIf iChars(i + 2&) = &H4C Then ' "L"
                                bArray(lCounter) = &H61: Skip = True
                            End If
                            
                        Case &H50 ' "P"
                            If iChars(i + 2&) = &H4B Then ' "K"
                                bArray(lCounter) = &H53: Skip = True
                            ElseIf iChars(i + 2&) = &H4F Then ' "O"
                                bArray(lCounter) = &H55: Skip = True
                            ElseIf iChars(i + 2&) = &H6F Then ' "o"
                                bArray(lCounter) = &H5E: Skip = True
                            End If
                            
                        Case &H42 ' "B"
                            If iChars(i + 2&) = &H4C Then ' "L"
                                bArray(lCounter) = &H57: Skip = True
                            End If
                        
                        Case &H4F ' "O"
                            If iChars(i + 2&) = &H43 Then ' "C"
                                bArray(lCounter) = &H58: Skip = True
                            End If
                    
                    End Select
                    
                    If Skip = True Then
                        i = i + 3&
                    End If
                    
                End If
            
            End If
            
            If Skip = False Then
            
                If i + 2& <= lUBound Then
                
                    If iChars(i + 2&) = &H5D Then ' "]"
                
                        Select Case iChars(i + 1&)
                            Case &H2E: bArray(lCounter) = &HB0: Skip = True ' "."
                            Case &H22: bArray(lCounter) = &HB1: Skip = True ' """
                            Case &H24: bArray(lCounter) = &HB7: Skip = True ' "$"
                            Case &H6D: bArray(lCounter) = &HB5: Skip = True ' "m"
                            Case &H66: bArray(lCounter) = &HB6: Skip = True ' "f"
                            Case &H4B: bArray(lCounter) = &H59: Skip = True ' "K"
                            Case &H41: bArray(lCounter) = &H62: Skip = True ' "A"
                            Case &H45: bArray(lCounter) = &H63: Skip = True ' "E"
                            Case &H55: bArray(lCounter) = &H79: Skip = True ' "U"
                            Case &H44: bArray(lCounter) = &H7A: Skip = True ' "D"
                            Case &H4C: bArray(lCounter) = &H7B: Skip = True ' "L"
                            Case &H52: bArray(lCounter) = &H7C: Skip = True ' "R"
                            Case &H27: bArray(lCounter) = &HB3: Skip = True ' "'"
                            Case &H78: bArray(lCounter) = &HB9: Skip = True ' "x"
                            Case &H3E: bArray(lCounter) = &HEF: Skip = True ' ">"
                            'Case &H75: bArray(lCounter) = &HF7: Skip = True ' "u"
                            'Case &H64: bArray(lCounter) = &HF8: Skip = True ' "d"
                            'Case &H6C: bArray(lCounter) = &HF9: Skip = True ' "l"
                        End Select
                
                        If Skip = True Then
                            i = i + 2&
                        End If
                    
                    End If
                
                End If
                
            End If
            
            Skip = False
            
        End If
        
        lCounter = lCounter + 1&
    
    Next i
    
    RtlMoveMemory ByVal lPointer, lSavePtr, 4&

End Sub

Public Function Asc2BrailleLen(ByVal sAscii As String) As Long
Dim iChars() As Integer
Dim Skip As Boolean
Dim lPointer As Long
Dim lSavePtr As Long
Dim lUBound As Long
Dim i As Long
    
    If Not AsciiReady Then
        AscInitialize
    End If
    
    ReDim iChars(1& To 1&) As Integer
    lSavePtr = VarPtr(iChars(1))
    
    RtlMoveMemory ByVal VarPtr(lPointer), ByVal ArrPtr(iChars), 4
    RtlMoveMemory ByVal lPointer + 16&, &H7FFFFFFF, 4
    lPointer = lPointer + 12&
    
    RtlMoveMemory ByVal lPointer, StrPtr(sAscii), 4
    lUBound = Len(sAscii)
  
    For i = LBound(iChars) To lUBound
        
        If iChars(i) = &H5C Then ' "\"
        
            If i + 1& <= lUBound Then
            
                Select Case iChars(i + 1&)
                        
                    Case &H4E: Skip = True ' "N"
                    Case &H58: Skip = True: Exit For ' "X"
                    
                    Case &H48 ' "H"
                    
                        If i + 3& <= lUBound Then
                    
                            If IsHexI(iChars(i + 2&), iChars(i + 3&)) Then
                                i = i + 3&
                            End If
                        
                        End If
                        
                End Select
                    
                If Skip = True Then
                    i = i + 1&
                End If
                
                Skip = False
            
            End If
            
        End If
        
        Asc2BrailleLen = Asc2BrailleLen + 1&
        
    Next i
    
    RtlMoveMemory ByVal lPointer, lSavePtr, 4
    
End Function

Public Sub Asc2Braille(ByVal sAscii As String, ByRef bArray() As Byte)
Dim iChars() As Integer
Dim Skip As Boolean
Dim lCounter As Long
Dim lPointer As Long
Dim lSavePtr As Long
Dim lUBound As Long
Dim i As Long
    
    If Not AsciiReady Then
        AscInitialize
    End If
    
    ReDim iChars(1& To 1&) As Integer
    lSavePtr = VarPtr(iChars(1))
    
    RtlMoveMemory ByVal VarPtr(lPointer), ByVal ArrPtr(iChars), 4
    RtlMoveMemory ByVal lPointer + 16&, &H7FFFFFFF, 4
    lPointer = lPointer + 12&
    
    RtlMoveMemory ByVal lPointer, StrPtr(sAscii), 4
    lUBound = Len(sAscii)
  
    For i = LBound(iChars) To lUBound
        
        If iChars(i) <> &H5C Then ' "\"
            bArray(lCounter) = bAsc2Braille(iChars(i))
        Else
            
            If i + 1& <= UBound(iChars) Then
            
                Select Case iChars(i + 1&)
                        
                    Case &H4E: bArray(lCounter) = &HFE: Skip = True ' "N"
                    Case &H58: bArray(lCounter) = &HFF: Skip = True: Exit For ' "X"
                    
                    Case &H48 ' "H"
                    
                        If i + 3& <= UBound(iChars) Then
                    
                            If IsHexI(iChars(i + 2&), iChars(i + 3&)) Then
                                bArray(lCounter) = CByte("&H" & ChrW$(iChars(i + 2&)) & ChrW$(iChars(i + 3&)))
                                i = i + 3&
                            End If
                        
                        End If
                        
                End Select
                    
                If Skip = True Then
                    i = i + 1&
                End If
                
                Skip = False
            
            End If
            
        End If

        lCounter = lCounter + 1
    
    Next i
    
    RtlMoveMemory ByVal lPointer, lSavePtr, 4
  
End Sub

Public Function Sapp2Asc(ByRef bSapp() As Byte, Optional Japanese As Boolean) As String
Dim lLength As Long
Dim i As Long

    If SappReady = False Then
        SappInitialize
    End If
    
    lLength = InStrB(bSapp, bEnd)
    
    If lLength <> 0 Then
        lLength = lLength - 1
    Else
        lLength = UBound(bSapp) + 1
    End If
    
    For i = LBound(bSapp) To lLength - 1

        If Not Japanese Then
            cString.Append sSapp2Asc(bSapp(i))
        Else
            cString.Append sSapp2AscJap(bSapp(i))
        End If
        
        If bSapp(i) > &HF7 Then
            
            Select Case bSapp(i)
                
                Case &HF8, &HF9, &HFC, &HFD
                
                    If i + 1 <= UBound(bSapp) Then
                        cString.Append "\h" & PadHex$(bSapp(i + 1), 2)
                        i = i + 1
                    End If
                
            End Select
            
        End If
    
    Next i
    
    Sapp2Asc = cString.ToString
    cString.Clear
  
    If InStrB(1, Sapp2Asc, "\v") Then

        DoReplace Sapp2Asc, "\v\h01", "[player]"
        DoReplace Sapp2Asc, "\v\h02", "[buffer1]"
        DoReplace Sapp2Asc, "\v\h03", "[buffer2]"
        DoReplace Sapp2Asc, "\v\h04", "[buffer3]"
        DoReplace Sapp2Asc, "\v\h06", "[rival]"
        DoReplace Sapp2Asc, "\v\h07", "[game]"
        DoReplace Sapp2Asc, "\v\h08", "[team]"
        DoReplace Sapp2Asc, "\v\h09", "[otherteam]"
        DoReplace Sapp2Asc, "\v\h0A", "[teamleader]"
        DoReplace Sapp2Asc, "\v\h0B", "[otherteamleader]"
        DoReplace Sapp2Asc, "\v\h0C", "[legend]"
        DoReplace Sapp2Asc, "\v\h0D", "[otherlegend]"
    
    End If
    
    If InStrB(1, Sapp2Asc, "\c") Then
    
        DoReplace Sapp2Asc, "\c\h06 ", "[small]"
        DoReplace Sapp2Asc, "\c\h15", "[jap]"
        DoReplace Sapp2Asc, "\c\h16", "[west]"
        
        Select Case sGameCode
        
            Case "AXV", "AXP"

                DoReplace Sapp2Asc, "\c\h01 ", "[transp_rs]"
                DoReplace Sapp2Asc, "\c\h01À", "[darkgrey_rs]"
                DoReplace Sapp2Asc, "\c\h01Á", "[red_rs]"
                DoReplace Sapp2Asc, "\c\h01Â", "[lightgreen_rs]"
                DoReplace Sapp2Asc, "\c\h01Ç", "[blue_rs]"
                DoReplace Sapp2Asc, "\c\h01È", "[yellow_rs]"
                DoReplace Sapp2Asc, "\c\h01É", "[cyan_rs]"
                DoReplace Sapp2Asc, "\c\h01Ê", "[magenta_rs]"
                DoReplace Sapp2Asc, "\c\h01Ë", "[grey_rs]"
                DoReplace Sapp2Asc, "\c\h01Ì", "[black_rs]"
                DoReplace Sapp2Asc, "\c\h0\h0A", "[black2_rs]"
                DoReplace Sapp2Asc, "\c\h01Î", "[lightgrey_rs]"
                DoReplace Sapp2Asc, "\c\h01Ï", "[white_rs]"
                DoReplace Sapp2Asc, "\c\h01Ò", "[skyblue_rs]"
                DoReplace Sapp2Asc, "\c\h01Ó", "[darkskyblue_rs]"
                DoReplace Sapp2Asc, "\c\h01Ô", "[white2_rs]"
                DoReplace Sapp2Asc, "\c\h01 ", "[black_rst]"
                DoReplace Sapp2Asc, "\c\h01À", "[white_rst]"
                DoReplace Sapp2Asc, "\c\h01Á", "[red_rst]"
                DoReplace Sapp2Asc, "\c\h01Â", "[navy_rst]"
                DoReplace Sapp2Asc, "\c\h01Ç", "[lightnavy_rst]"
                DoReplace Sapp2Asc, "\c\h01È", "[white2_rst]"
                DoReplace Sapp2Asc, "\c\h01É", "[darkpurple_rst]"
                DoReplace Sapp2Asc, "\c\h01Ê", "[lightpurple_rst]"
                DoReplace Sapp2Asc, "\c\h01Ë", "[darknavy_rst]"
                DoReplace Sapp2Asc, "\c\h01Ì", "[darkgrey_rst]"
                DoReplace Sapp2Asc, "\c\h01\h0A", "[grey_rst]"
                DoReplace Sapp2Asc, "\c\h01Î", "[darkbronze_rst]"
                DoReplace Sapp2Asc, "\c\h01Ï", "[bronze_rst]"
                DoReplace Sapp2Asc, "\c\h01Ò", "[darkred_rst]"
                DoReplace Sapp2Asc, "\c\h01Ó", "[purple_rst]"
                DoReplace Sapp2Asc, "\c\h01Ô", "[transp_rst]"
        
            Case "BPR", "BPG"
                
                DoReplace Sapp2Asc, "\c\h01 ", "[white_fr]"
                DoReplace Sapp2Asc, "\c\h01À", "[white2_fr]"
                DoReplace Sapp2Asc, "\c\h01Á", "[black_fr]"
                DoReplace Sapp2Asc, "\c\h01Â", "[grey_fr]"
                DoReplace Sapp2Asc, "\c\h01Ç", "[red_fr]"
                DoReplace Sapp2Asc, "\c\h01È", "[orange_fr]"
                DoReplace Sapp2Asc, "\c\h01É", "[green_fr]"
                DoReplace Sapp2Asc, "\c\h01Ê", "[lightgreen_fr]"
                DoReplace Sapp2Asc, "\c\h01Ë", "[blue_fr]"
                DoReplace Sapp2Asc, "\c\h01Ì", "[lightblue_fr]"
                DoReplace Sapp2Asc, "\c\h01\h0A", "[white3_fr]"
                DoReplace Sapp2Asc, "\c\h01Î", "[lightblue2_fr]"
                DoReplace Sapp2Asc, "\c\h01Ï", "[cyan_fr]"
                DoReplace Sapp2Asc, "\c\h01Ò", "[lightblue3_fr]"
                DoReplace Sapp2Asc, "\c\h01Ó", "[navyblue_fr]"
                DoReplace Sapp2Asc, "\c\h01Ô", "[darknavyblue_fr]"
                DoReplace Sapp2Asc, "\c\h01 ", "[darknavy_frt]"
                DoReplace Sapp2Asc, "\c\h01À", "[white_frt]"
                DoReplace Sapp2Asc, "\c\h01Á", "[red_frt]"
                DoReplace Sapp2Asc, "\c\h01Â", "[navy_frt]"
                DoReplace Sapp2Asc, "\c\h01Ç", "[lightnavy_frt]"
                DoReplace Sapp2Asc, "\c\h01È", "[white_frt]"
                DoReplace Sapp2Asc, "\c\h01É", "[purple2_frt]"
                DoReplace Sapp2Asc, "\c\h01Ê", "[grey_frt]"
                DoReplace Sapp2Asc, "\c\h01Ë", "[darkpurple_frt]"
                DoReplace Sapp2Asc, "\c\h01Ì", "[black_frt]"
                DoReplace Sapp2Asc, "\c\h01\h0A", "[magenta_frt]"
                DoReplace Sapp2Asc, "\c\h01Î", "[white2_frt]"
                DoReplace Sapp2Asc, "\c\h01Ï", "[magenta_frt]"
                DoReplace Sapp2Asc, "\c\h01Ò", "[gold_frt]"
                DoReplace Sapp2Asc, "\c\h01Ó", "[lightgold_frt]"
                DoReplace Sapp2Asc, "\c\h01Ô", "[darknavy2_frt]"
                
            Case "BPE"
        
                DoReplace Sapp2Asc, "\c\h01 ", "[white_em]"
                DoReplace Sapp2Asc, "\c\h01À", "[white2_em]"
                DoReplace Sapp2Asc, "\c\h01Á", "[darkgrey_em]"
                DoReplace Sapp2Asc, "\c\h01Â", "[grey_em]"
                DoReplace Sapp2Asc, "\c\h01Ç", "[red_em]"
                DoReplace Sapp2Asc, "\c\h01È", "[orange_em]"
                DoReplace Sapp2Asc, "\c\h01É", "[green_em]"
                DoReplace Sapp2Asc, "\c\h01Ê", "[lightgreen_em]"
                DoReplace Sapp2Asc, "\c\h01Ë", "[blue_em]"
                DoReplace Sapp2Asc, "\c\h01Ì", "[lightblue_em]"
                DoReplace Sapp2Asc, "\c\h01\h0A", "[white3_em]"
                DoReplace Sapp2Asc, "\c\h01Î", "[white4_em]"
                DoReplace Sapp2Asc, "\c\h01Ï", "[white5_em]"
                DoReplace Sapp2Asc, "\c\h01Ò", "[limegreen_em]"
                DoReplace Sapp2Asc, "\c\h01Ó", "[aqua_em]"
                DoReplace Sapp2Asc, "\c\h01Ô", "[navy_em]"
                DoReplace Sapp2Asc, "\c\h01 ", "[transp_emt]"
                DoReplace Sapp2Asc, "\c\h01À", "[white_emt]"
                DoReplace Sapp2Asc, "\c\h01Á", "[red_emt]"
                DoReplace Sapp2Asc, "\c\h01Â", "[darknavy_emt]"
                DoReplace Sapp2Asc, "\c\h01Ç", "[navy_emt]"
                DoReplace Sapp2Asc, "\c\h01È", "[white2_emt]"
                DoReplace Sapp2Asc, "\c\h01É", "[grey_emt]"
                DoReplace Sapp2Asc, "\c\h01Ê", "[lightgrey_emt]"
                DoReplace Sapp2Asc, "\c\h01Ë", "[darknavy2_emt]"
                DoReplace Sapp2Asc, "\c\h01Ì", "[darkgrey_emt]"
                DoReplace Sapp2Asc, "\c\h01\h0A", "[lightgrey_emt]"
                DoReplace Sapp2Asc, "\c\h01Î", "[brown_emt]"
                DoReplace Sapp2Asc, "\c\h01Ï", "[lightbrown_emt]"
                DoReplace Sapp2Asc, "\c\h01Ò", "[darkred_emt]"
                DoReplace Sapp2Asc, "\c\h01Ó", "[purple_emt]"
                DoReplace Sapp2Asc, "\c\h01Ô", "[transp2_emt]"
            
        End Select
    
    End If

End Function

Public Function Braille2Asc(ByRef bBraille() As Byte) As String
Dim lLength As Long
Dim i As Long
    
    If BrailleReady = False Then
        BrailleInitialize
    End If
    
    lLength = InStrB(bBraille, bEnd)
    
    If lLength <> 0 Then
        lLength = lLength - 1
    Else
        lLength = UBound(bBraille) + 1
    End If
    
    For i = LBound(bBraille) To lLength - 1
        cString.Append sBraille2Asc(bBraille(i))
    Next i
    
    Braille2Asc = cString.ToString
    cString.Clear

End Function
