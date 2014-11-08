Attribute VB_Name = "modAssociate"
Option Explicit

'Microsoft's answers to associating files are:
'1. HOWTO: Associate a File Extension with Your Application
'http://support.microsoft.com/default.aspx?scid=KB;en-us;q185453
'
'2. HOWTO: Associate a Custom Icon with a File Extension
'http://support.microsoft.com/default.aspx?scid=kb;en-us;247529

'========Read regisy key values
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

'Note that if you declare the lpData parameter as String,
'you must pass it By Value. (In RegQueryValueEx)
Public phkResult As Long
Public lpSubKey As String
Public lpData As String
Public lpcbData As Long
'Public RC As Long

'Root Key Constants ...................................
Private Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const sUserClasses As String = "Software\Classes\"
'Reg DataType Constants ...............................
Public Const REG_SZ = 1 ' Unicode null terminated string

'===============Create and delete key in regisy
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
      
      ' Return codes from Registration functions.
'      Const ERROR_SUCCESS = 0&
'      Const ERROR_BADDB = 1&
'      Const ERROR_BADKEY = 2&
'      Const ERROR_CANTOPEN = 3&
'      Const ERROR_CANTREAD = 4&
'      Const ERROR_CANTWRITE = 5&
'      Const ERROR_OUTOFMEMORY = 6&
'      Const ERROR_INVALID_PARAMETER = 7&
'      Const ERROR_ACCESS_DENIED = 8&
      
    Private Const MAX_PATH = 260&
      
'This sub puts new default icon on associated files or off if unassociated
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0&
Private Const SHCNF_FLUSHNOWAIT As Long = &H2000

Public Sub IntegrateShell(Extension As String, PathToExecute As String, CommandName As String)
Dim sKeyName As String   'Holds Key Name in regisy.
Dim sKeyValue As String  'Holds Key Value in regisy.
Dim lphKey&        'Holds created key handle from RegCreateKey.
Dim sTemp As String
Dim lRet As Long

    sKeyName = "." & Extension
    sKeyValue = ChrW$(34) & PathToExecute & ChrW$(34) & " " & ChrW$(34) & "%1" & ChrW$(34)
    
    RegOpenKey HKEY_CURRENT_USER, sUserClasses & sKeyName, phkResult
    
    lpData = Space$(255)
    lpcbData = 255
    lRet = RegQueryValueEx(phkResult, vbNullString, 0&, REG_SZ, lpData, lpcbData)
    sTemp = Mid$(lpData, 1, lpcbData - 1)
    
    RegCreateKey& HKEY_CURRENT_USER, sUserClasses & sKeyName, lphKey&
    RegSetValue& lphKey&, "shell\" & CommandName, REG_SZ, LoadResString(4014), MAX_PATH
    RegSetValue& lphKey&, "shell\" & CommandName & "\command", REG_SZ, sKeyValue, MAX_PATH
    
    If lRet = 0 Then
        If LenB(sTemp) <> 0 Then
            RegCreateKey& HKEY_CURRENT_USER, sUserClasses & sTemp, lphKey&
            RegSetValue& lphKey&, "shell\" & CommandName, REG_SZ, LoadResString(4014), MAX_PATH
            RegSetValue& lphKey&, "shell\" & CommandName & "\command", REG_SZ, sKeyValue, MAX_PATH
        End If
    End If
    
End Sub

Public Sub UnintegrateShell(Extension As String, CommandName As String)
Dim sKeyName As String   'Holds Key Name in regisy.
Dim sTemp As String
Dim lRet As Long
   
    sKeyName = "." & Extension
    
    RegOpenKey HKEY_CURRENT_USER, sUserClasses & sKeyName, phkResult
    
    lpData = Space$(255)
    lpcbData = 255
    lRet = RegQueryValueEx(phkResult, vbNullString, 0&, REG_SZ, lpData, lpcbData)
    sTemp = Mid$(lpData, 1, lpcbData - 1)
    
    RegDeleteKey HKEY_CURRENT_USER, sUserClasses & sKeyName & "\shell\" & CommandName & "\command"
    RegDeleteKey HKEY_CURRENT_USER, sUserClasses & sKeyName & "\shell\" & CommandName
    RegDeleteKey HKEY_CLASSES_ROOT, sKeyName & "\shell\" & CommandName & "\command"
    RegDeleteKey HKEY_CLASSES_ROOT, sKeyName & "\shell\" & CommandName
    
    If lRet = 0 Then
        If LenB(sTemp) <> 0 Then
            sKeyName = sTemp
            RegDeleteKey HKEY_CURRENT_USER, sUserClasses & sKeyName & "\shell\" & CommandName & "\command"
            RegDeleteKey HKEY_CURRENT_USER, sUserClasses & sKeyName & "\shell\" & CommandName
            RegDeleteKey HKEY_CLASSES_ROOT, sKeyName & "\shell\" & CommandName & "\command"
            RegDeleteKey HKEY_CLASSES_ROOT, sKeyName & "\shell\" & CommandName
        End If
    End If

End Sub

'Extension is three letters without the "."
'PathToExecute is full path to exe file
'Application Name is any name you want as description of Extension

Public Sub AssociateExt(Extension As String, PathToExecute As String, ApplicationName As String, Optional iIconIndex As Integer = 0)
Dim sKeyName As String   'Holds Key Name in regisy.
Dim sKeyValue As String  'Holds Key Value in regisy.
Dim lphKey&        'Holds created key handle from RegCreateKey.

    If InStrB(1, Extension, ".") <> 0 Then
        'MsgBox "Extension has . in it. Remove and try again."
        Exit Sub
    End If
    
    'This creates a Root entry for the extension to be associated with 'ApplicationName'.
    sKeyName = "." & Extension
    sKeyValue = ApplicationName
    RegCreateKey& HKEY_CURRENT_USER, sUserClasses & sKeyName, lphKey&
    RegSetValue& lphKey&, vbNullString, REG_SZ, sKeyValue, 0&
      
    'This creates a Root entry called 'ApplicationName'.
    sKeyName = ApplicationName
    sKeyValue = sKeyName
    RegCreateKey& HKEY_CURRENT_USER, sUserClasses & sKeyName, lphKey&
    RegSetValue& lphKey&, vbNullString, REG_SZ, sKeyValue, 0&

    'This sets the command line for 'ApplicationName'.
    'sKeyName = ApplicationName
    sKeyValue = ChrW$(34) & PathToExecute & ChrW$(34) & " " & ChrW$(34) & "%1" & ChrW$(34)
    RegCreateKey& HKEY_CURRENT_USER, sUserClasses & sKeyName, lphKey&
    RegSetValue& lphKey&, "shell\open\command", REG_SZ, sKeyValue, MAX_PATH

    'This sets the default icon
    'sKeyName = ApplicationName
    sKeyValue = App.Path & "\" & App.EXEName & ".exe," & iIconIndex
    RegCreateKey& HKEY_CURRENT_USER, sUserClasses & sKeyName, lphKey&
    RegSetValue& lphKey&, "DefaultIcon", REG_SZ, sKeyValue, MAX_PATH

    'Force Icon Refresh
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
    
End Sub

Public Sub UnAssociateExt(Extension As String, ApplicationName As String)
Dim sKeyName As String   'Finds Key Name in regisy.
'Dim sKeyValue As String  'Finds Key Value in regisy.

    If InStrB(1, Extension, ".") <> 0 Then
        'MsgBox "Extension has . in it. Remove and try again."
        Exit Sub
    End If

    'This deletes the default icon
    sKeyName = ApplicationName
    RegDeleteKey HKEY_CURRENT_USER, sUserClasses & sKeyName & "\DefaultIcon"
    RegDeleteKey HKEY_CLASSES_ROOT, sKeyName & "\DefaultIcon"

    'This deletes the command line for "ApplicationName".
    'sKeyName = ApplicationName
    RegDeleteKey HKEY_CURRENT_USER, sUserClasses & sKeyName & "\shell\open\command"
    RegDeleteKey HKEY_CURRENT_USER, sUserClasses & sKeyName & "\shell\open"
    RegDeleteKey HKEY_CURRENT_USER, sUserClasses & sKeyName & "\shell"
    RegDeleteKey HKEY_CURRENT_USER, sUserClasses & sKeyName
    RegDeleteKey HKEY_CLASSES_ROOT, sKeyName & "\shell\open\command"
    RegDeleteKey HKEY_CLASSES_ROOT, sKeyName & "\shell\open"
    RegDeleteKey HKEY_CLASSES_ROOT, sKeyName & "\shell"
    RegDeleteKey HKEY_CLASSES_ROOT, sKeyName

    'This deletes the Root entry for the extension to be associated with "ApplicationName".
    sKeyName = "." & Extension
    RegDeleteKey HKEY_CURRENT_USER, sUserClasses & sKeyName
    RegDeleteKey HKEY_CLASSES_ROOT, sKeyName

    'Force Icon Refresh
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST & SHCNF_FLUSHNOWAIT, 0, 0
    'Thanks to Ralf Gerstenberger <ralf.gerstenberger@arcor.de> for pointing out
    'that WinXP seems to require the SHCNF_FLUSHNOWAIT flag in SHChangeNotify
    'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/functions/shchangenotify.asp
    
End Sub
