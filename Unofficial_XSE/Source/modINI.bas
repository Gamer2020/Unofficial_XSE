Attribute VB_Name = "modINI"
Option Explicit

Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal sSectionName As String, ByVal sString As String, ByVal sFileName As String) As Long

Private Const MAX_STRING_LEN As Integer = 260
Private Const MAX_SECTION_LEN As Long = 65536

Public Const IniFile As String = "\Settings.ini"
    
'Read a string from an ini file
Public Function ReadIniString(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, Optional ByVal Default As String = vbNullString) As String
Dim nLength As Long
Dim sTemp As String

    On Error GoTo ReadIniString_err
    
    sTemp = Space$(MAX_STRING_LEN)
    nLength = GetPrivateProfileString(Section, Key, Default, sTemp, MAX_STRING_LEN, IniFile)
    
    If nLength > 0 Then
        ReadIniString = Left$(sTemp, nLength)
    End If
    
    Exit Function
    
ReadIniString_err:
    ReadIniString = Default
End Function

'Read a long integer from an ini file
Public Function ReadIniLong(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, Optional ByVal Default As Long = 0) As Long
Dim sTemp As String

    On Error GoTo ReadIniLong_err
    
    'Use existing function to get value back as string
    sTemp = ReadIniString(IniFile, Section, Key, "")

    If LenB(sTemp) = 0 Then
        ReadIniLong = Default
    Else
        ReadIniLong = CLng(sTemp)
    End If
    
    Exit Function
    
ReadIniLong_err:
    ReadIniLong = Default
End Function

'Read a double from an ini file
'Public Function dReadIniFileDouble(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, Optional ByVal Default As Double = 0) As Double
'    On Error GoTo dReadIniFileDouble_err
'    Dim sTemp As String
'    'Use existing function to get value back
'    '     as string
'    sTemp = ReadIniString(IniFile, Section, Key, "")
'
'    If Len(sTemp) = 0 Then
'        dReadIniFileDouble = Default
'    Else
'        dReadIniFileDouble = CDbl(sTemp)
'    End If
'    Exit Function
'dReadIniFileDouble_err:
'    dReadIniFileDouble = Default
'End Function

'This will return a collection containing all entries for a given section
Public Function ReadIniSection(ByVal IniFile As String, ByVal Section As String) As Collection
Dim sTemp As String
Dim nPos As Long
Dim nLength As Long
    
    On Error GoTo ReadIniSection_err
    Set ReadIniSection = New Collection
    
    sTemp = Space$(MAX_SECTION_LEN)
    nLength = GetPrivateProfileSection(Section, sTemp, MAX_SECTION_LEN, IniFile)
    
    If nLength > 0 Then
    
       sTemp = Left$(sTemp, nLength)
       nPos = InStr(sTemp, "=")
    
       Do While nPos > 0
           ReadIniSection.Add Mid$(sTemp, 1, nPos - 1)
           nPos = InStr(sTemp, vbNullChar)
           sTemp = Mid$(sTemp, nPos + 1)
           nPos = InStr(sTemp, "=")
       Loop
    
       If LenB(sTemp) <> 0 Then
           ReadIniSection.Add sTemp
       End If
    
    End If
        
    Exit Function
    
ReadIniSection_err:
End Function

'Write a string to an ini file
Public Function WriteStringToIni(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
Dim nRetVal As Long
    
    'On Error GoTo WriteStringToIni_err
    MakeWritable IniFile
    
    'Clear existing entry first
    '(there is a problem with encrypted values otherwise)
    WritePrivateProfileString Section, Key, vbNullString, IniFile
    nRetVal = WritePrivateProfileString(Section, Key, Value, IniFile)
 
    If nRetVal > 0 Then
        WriteStringToIni = True
    Else
        WriteStringToIni = False
    End If
    
    Exit Function
    
WriteStringToIni_err:
    WriteStringToIni = False
End Function

'Write a long integer to an ini file
'Public Function bWriteIniFileLong(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As Long) As Boolean
'    bWriteIniFileLong = WriteStringToIni(IniFile, Section, Key, CStr(Value))
'End Function

'Write a double to an ini file
'Public Function bWriteIniFileDouble(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As Double) As Boolean
'    bWriteIniFileDouble = WriteStringToIni(IniFile, Section, Key, CStr(Value))
'End Function

'Write a date to an ini file
'Public Function bWriteIniFileDate(ByVal IniFile As String, ByVal Section As String, ByVal Key As String, ByVal Value As Date) As Boolean
'    bWriteIniFileDate = WriteStringToIni(IniFile, Section, Key, Format$(Value, "dd mmm yyyy hh:nn:ss"))
'End Function

'This will return a collection containing all entries for a given section
'Public Function ReadIniSectionNames(ByVal INIFile As String) As Collection
'    On Error GoTo ReadIniSectionNames_err
'    Dim sTemp As String
'    Dim nPos As Long
'    Dim nLength As Long
'    Set ReadIniSectionNames = New Collection
'    sTemp = SysAllocStringLen(vbNullString, MAX_SECTION_LEN)
'    nLength = GetPrivateProfileSectionNames(sTemp, MAX_SECTION_LEN, INIFile)
'    sTemp = Mid$(sTemp, 1, nLength)
'    nPos = InStr(1, sTemp, vbNullChar)
'
'    Do While nPos > 0
'        ReadIniSectionNames.Add Mid$(sTemp, 1, nPos - 1)
'        sTemp = Mid$(sTemp, nPos + 1)
'        nPos = InStr(1, sTemp, vbNullChar)
'
'        DoEvents
'        Loop
'
'        If Len(sTemp) > 0 Then
'            ReadIniSectionNames.Add sTemp
'        End If
'        Exit Function
'ReadIniSectionNames_err:
'    End Function
    
'This will remove an entry
'Public Function bRemoveIniFileEntry(ByVal INIFile As String, ByVal Section As String, ByVal Key As String) As Boolean
'    On Error GoTo bRemoveIniFileEntry_err
'    bRemoveIniFileEntry = False
'
'    SetAttrIfExists INIFile, vbNormal
'
'    If WritePrivateProfileString(Section, Key, vbNullString, INIFile) > 0 Then
'        bRemoveIniFileEntry = True
'    Else
'        bRemoveIniFileEntry = False
'    End If
'bRemoveIniFileEntry_err:
'End Function

'This will remove a section
Public Function RemoveIniSection(ByVal IniFile As String, ByVal Section As String) As Boolean
    
    On Error GoTo RemoveIniSection_err
    MakeWritable IniFile

    If WritePrivateProfileSection(Section, String$(2, vbNullChar), IniFile) > 0 Then
        RemoveIniSection = True
    End If
    
RemoveIniSection_err:
End Function

'Rename a string in an INI file
'Public Function bRenameIniFileString(ByVal INIFile As String, ByVal Section As String, ByVal Key As String, ByVal NewKey As String) As Boolean
'    On Error GoTo bRenameIniFileEntry_err
'
'    Dim tmpValue As String
'    tmpValue = ReadIniString(INIFile, Section, Key) ' store the previous key value
'    bRemoveIniFileEntry INIFile, Section, Key ' remove old key
'    WriteStringToIni INIFile, Section, NewKey, tmpValue ' create the new renamed one
'
'bRenameIniFileEntry_err:
'End Function
