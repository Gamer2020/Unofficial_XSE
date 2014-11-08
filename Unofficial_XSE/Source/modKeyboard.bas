Attribute VB_Name = "modKeyboard"
Option Explicit

Private Const VER_PLATFORM_WIN32_WINDOWS = 1

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2

Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState2 Lib "user32" Alias "GetKeyboardState" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbkeystate As Byte) As Long

Public Sub ToggleKey(bKey As Byte, TurnOn As Boolean)
Dim bKeys(255) As Byte
Dim bKeyOn As Boolean
Dim typOS As OSVERSIONINFO

      'Get status of the 256 virtual keys
      GetKeyboardState2 bKeys(0)
      bKeyOn = bKeys(bKey)
      
      If bKeyOn <> TurnOn Then
        If typOS.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then ' Win95/98
          bKeys(bKey) = 1
          SetKeyboardState bKeys(0)
        Else  'WinNT/2000

        'Simulate Key Press
          keybd_event bKey, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
          keybd_event bKey, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        End If
      End If
     
End Sub

Public Function CapsLockOn() As Boolean
Dim iKeyState As Integer
    
    iKeyState = GetKeyState(vbKeyCapital)
    CapsLockOn = (iKeyState = 1 Or iKeyState = -127)
    
End Function

Public Function NumLockOn() As Boolean
Dim iKeyState As Integer
    
    iKeyState = GetKeyState(vbKeyNumlock)
    NumLockOn = (iKeyState = 1 Or iKeyState = -127)
    
End Function

Public Function ScrollLockOn() As Boolean
Dim iKeyState As Integer
    
    iKeyState = GetKeyState(vbKeyScrollLock)
    ScrollLockOn = (iKeyState = 1 Or iKeyState = -127)
    
End Function

Public Sub GetKeyStatus()
    frmMain.StatusBar.PanelEnabled(5) = CapsLockOn
    frmMain.StatusBar.PanelEnabled(6) = NumLockOn
    frmMain.StatusBar.PanelEnabled(7) = ScrollLockOn
End Sub
