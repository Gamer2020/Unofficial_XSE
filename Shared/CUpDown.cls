VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©2002 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

' Win32 API Declares
Private Declare Function CreateUpDownControl Lib "COMCTL32.DLL" (ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal hParent As Long, ByVal nID As Long, ByVal hInst As Long, ByVal hBuddy As Long, ByVal nUpper As Long, ByVal nLower As Long, ByVal nPos As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)

' Up/Down Styles
Private Const UDS_ALIGNLEFT As Long = &H8
Private Const UDS_ALIGNRIGHT As Long = &H4
Private Const UDS_ARROWKEYS As Long = &H20
'Private Const UDS_AUTOBUDDY As Long = &H10
Private Const UDS_HORZ As Long = &H40
Private Const UDS_HOTTRACK As Long = &H100
Private Const UDS_NOTHOUSANDS As Long = &H80
Private Const UDS_SETBUDDYINT As Long = &H2
'Private Const UDS_WRAP As Long = &H1

' Up/Down 16-bit Min/Max Values.
Private Const UD_MAXVAL As Long = &H7FFF
Private Const UD_MINVAL As Long = (-UD_MAXVAL)

' Up/Down Messages
Private Const WM_USER As Long = &H400
'Private Const CCM_FIRST As Long = &H2000
'Private Const CCM_GETUNICODEFORMAT As Long = (CCM_FIRST + 6)
'Private Const CCM_SETUNICODEFORMAT As Long = (CCM_FIRST + 5)
'Private Const UDM_GETACCEL As Long = (WM_USER + 108)
'Private Const UDM_GETBASE As Long = (WM_USER + 110)
'Private Const UDM_GETBUDDY As Long = (WM_USER + 106)
Private Const UDM_GETPOS As Long = (WM_USER + 104)
'Private Const UDM_GETPOS32 As Long = (WM_USER + 114)
Private Const UDM_GETRANGE As Long = (WM_USER + 102)
'Private Const UDM_GETRANGE32 As Long = (WM_USER + 112)
'Private Const UDM_GETUNICODEFORMAT As Long = CCM_GETUNICODEFORMAT
Private Const UDM_SETACCEL As Long = (WM_USER + 107)
Private Const UDM_SETBASE As Long = (WM_USER + 109)
'Private Const UDM_SETBUDDY As Long = (WM_USER + 105)
Private Const UDM_SETPOS As Long = (WM_USER + 103)
'Private Const UDM_SETPOS32 As Long = (WM_USER + 113)
Private Const UDM_SETRANGE As Long = (WM_USER + 101)
'Private Const UDM_SETRANGE32 As Long = (WM_USER + 111)
'Private Const UDM_SETUNICODEFORMAT As Long = CCM_SETUNICODEFORMAT

' Structure used for acceleration.
Private Type UDACCEL
   nSec As Long
   nInc As Long
End Type

' Structure useds to determine window size/position.
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type POINTAPI
   X As Long
   Y As Long
End Type

' Window Styles
Private Const WS_BORDER   As Long = &H800000
Private Const WS_DISABLED As Long = &H8000000
Private Const WS_CHILD    As Long = &H40000000
Private Const WS_VISIBLE  As Long = &H10000000

' Used to get window style bits.
'Private Const GWL_STYLE = (-16)
'Private Const GWL_EXSTYLE = (-20)

' Force total redraw that shows new styles.
'Private Const SWP_FRAMECHANGED = &H20
'Private Const SWP_NOMOVE = &H2
'Private Const SWP_NOZORDER = &H4
'Private Const SWP_NOSIZE = &H1

' ShowWindow() Commands
'Private Const SW_HIDE = 0
'Private Const SW_SHOWNORMAL = 1
'Private Const SW_NORMAL = 1
'Private Const SW_SHOWMINIMIZED = 2
'Private Const SW_SHOWMAXIMIZED = 3
'Private Const SW_MAXIMIZE = 3
'Private Const SW_SHOWNOACTIVATE = 4
'Private Const SW_SHOW = 5
'Private Const SW_MINIMIZE = 6
'Private Const SW_SHOWMINNOACTIVE = 7
'Private Const SW_SHOWNA = 8
'Private Const SW_RESTORE = 9
'Private Const SW_SHOWDEFAULT = 10
'Private Const SW_MAX = 10

' Used to refresh buddy display.
Private Const WM_SETTEXT As Long = &HC

' Member variables
Private m_hWnd As Long
Private WithEvents m_Buddy As TextBox
Attribute m_Buddy.VB_VarHelpID = -1
Private m_hBuddy As Long
Private m_OrigRect As RECT
Private m_Acceleration As udAccelerations
Private m_Align As udAlignments
Private m_ArrowKeys As Long
Private m_Base As Long
Private m_Border As Boolean
Private m_Direction As udDirections
Private m_Enabled As Boolean
Private m_HotTrack As Long
'Private m_Increment As Long
Private m_Max As Long
Private m_Min As Long
Private m_Thousands As Long
Private m_Value As Long
Private m_Visible As Boolean
Private m_Wrap As Long

' Defaults
Private Const defAcceleration As Long = 2 'udAccelerateMedium
Private Const defAlignment As Long = UDS_ALIGNRIGHT
Private Const defArrowKeys As Long = UDS_ARROWKEYS
Private Const defBase As Long = 10 'udBase10
Private Const defBorder As Boolean = False
Private Const defDirection As Long = 0 'udVertical
Private Const defEnabled As Boolean = False
Private Const defHotTrack As Long = UDS_HOTTRACK
'Private Const defIncrement As Long = 1
Private Const defThousands As Long = 0 'Not(UDS_NOTHOUSANDS)
Private Const defVisible As Boolean = True
Private Const defWrap As Long = 0 'Not(UDS_WRAP)

' Enumerations
Public Enum udAccelerations
   udAccelerateNone = 0
   udAccelerateSlow = 1
   udAccelerateMedium = 2
   udAccelerateFast = 3
End Enum

Public Enum udAlignments
   udAlignLeft = UDS_ALIGNLEFT
   udAlignRight = UDS_ALIGNRIGHT
End Enum

Public Enum udDirections
   udVertical = 0
   udHorizontal = UDS_HORZ
End Enum

Public Enum udRadixBases
   udBase10 = 10
   udbase16 = 16
End Enum

' ***********************************************
'  Initialization and Termination
' ***********************************************
Private Sub Class_Initialize()
   ' Default values.
   m_Acceleration = defAcceleration
   m_Align = defAlignment
   m_ArrowKeys = defArrowKeys
   m_Base = defBase
   m_Border = defBorder
   m_Direction = defDirection
   m_Enabled = defEnabled
   m_HotTrack = defHotTrack
'   m_Increment = defIncrement
   m_Thousands = defThousands
   m_Visible = defVisible
   m_Wrap = defWrap
End Sub

Private Sub Class_Terminate()
   ' Kill existing up/down, if it exists,
   ' restoring original position/size of
   ' buddy textbox.
   If IsWindow(m_hWnd) Then
      Call DestroyWindow(m_hWnd)
      With m_OrigRect
         Call MoveWindow(m_hBuddy, .Left, .Top, .Right, .Bottom, True)
      End With
   End If
End Sub

' ***********************************************
'  Buddy Events
' ***********************************************
' TODO: Update to use UDM_GETPOS32 if IE5+ is installed.
Private Sub m_Buddy_Change()
   ' We cannot sink events from any control that
   ' is a member of a control array, so we need
   ' to provide a public method to manually check
   ' the user's input.
   Me.Validate
   
   If LenB(m_Buddy.Text) <> 0 Then
     If Enabled = False Then
        Enabled = True
     End If
   End If
   
End Sub

' ***********************************************
'  Public Properties -- ReadOnly!
' ***********************************************
'Public Property Get hWnd() As Long
'   ' Return handle to updown control.
'   hWnd = m_hWnd
'End Property

' ***********************************************
'  Public Properties -- General
' ***********************************************
'Public Property Let Acceleration(ByVal NewVal As udAccelerations)
'   If NewVal >= udAccelerateNone And NewVal <= udAccelerateFast Then
'      ' Cache value, and attempt to set.
'      If m_Acceleration <> NewVal Then
'         m_Acceleration = NewVal
'         Call AccelerationSet
'      End If
'   End If
'End Property

'Public Property Get Acceleration() As udAccelerations
'   ' Return cached value.
'   Acceleration = m_Acceleration
'End Property

'Public Property Let Alignment(ByVal NewVal As udAlignments)
'   ' Validate and cache.
'   If NewVal = udAlignLeft Or NewVal = udAlignRight Then
'      ' No need to cache/recreate unless this is a new value.
'      If m_Align <> NewVal Then
'         m_Align = NewVal
'         ' This style can only be set when control is created.
'         If m_hWnd Then Recreate
'      End If
'   End If
'End Property

'Public Property Get Alignment() As udAlignments
'   ' Return cached value.
'   Alignment = m_Align
'End Property

'Public Property Let ArrowKeys(ByVal NewVal As Boolean)
'   Dim NewStyle As Long
'
'   ' Interpret style bit to use for this setting.
'   If NewVal = True Then
'      NewStyle = UDS_ARROWKEYS
'   Else
'      NewStyle = 0&
'   End If
'
'   ' No need to cache/recreate unless this is a new value.
'   If m_ArrowKeys <> NewStyle Then
'      m_ArrowKeys = NewStyle
'      ' This style can only be set when control is created.
'         If m_hWnd Then Recreate
'   End If
'End Property

'Public Property Get ArrowKeys() As Boolean
'   ' Return cached value.
'   ArrowKeys = CBool(m_ArrowKeys)
'End Property

Public Property Let Base(ByVal NewVal As udRadixBases)
   ' Validate and cache.
   If (NewVal = udBase10) Or (NewVal = udbase16) Then
      ' No need to cache/recreate unless this is a new value.
      If m_Base <> NewVal Then
         m_Base = NewVal
         ' We can only set this after window creation.
         Call SendMessage(m_hWnd, UDM_SETBASE, m_Base, ByVal 0&)
         ' Values less than 0x000A are not "properly"
         ' displayed without a refresh, when switching
         ' from base10 to base16.
         If m_Value >= 0 And m_Value < &HA Then
            If m_Base = udbase16 Then
               Me.Refresh
            End If
         End If
      End If
   End If
End Property

Public Property Get Base() As udRadixBases
   ' Return cached value.
   Base = m_Base
End Property

'Public Property Let Border(ByVal NewVal As Boolean)
'   ' No need to do anything unless this is a new value.
'   If m_Border <> NewVal Then
'      m_Border = NewVal
'      ' Update style on the fly, if window exists.
'      If m_hWnd Then
'         Call FlipBit(WS_BORDER, m_Border)
'      End If
'   End If
'End Property
   
'Public Property Get Border() As Boolean
'   ' Check and cache actual value, then return.
'   m_Border = CBool(Style() And WS_BORDER)
'   Border = m_Border
'End Property

Public Property Set Buddy(ByVal TxtBox As TextBox)
   ' We have a buddy! Time to (re)create up/down.
   m_hBuddy = TxtBox.hwnd
   Call Recreate
   Call SendMessage(m_hBuddy, WM_SETTEXT, 0&, ByVal "")
   ' Attempt to cache a reference.  This will fail if
   ' the textbox belongs to a control array, as we have
   ' declared this using WithEvents to capture _Change.
   On Error Resume Next
   Set m_Buddy = TxtBox
End Property

Public Property Get Buddy() As TextBox
   ' Return cached object.
   Set Buddy = m_Buddy
End Property

'Public Property Let Direction(ByVal NewVal As udDirections)
'   ' Validate and cache.
'   If (NewVal = udHorizontal) Or (NewVal = udVertical) Then
'      ' No need to cache/recreate unless this is a new value.
'      If m_Direction <> NewVal Then
'         m_Direction = NewVal
'         ' This style can only be set when control is created.
'         If m_hWnd Then Recreate
'      End If
'   End If
'End Property

'Public Property Get Direction() As udDirections
'   ' Return cached value.
'   Direction = m_Direction
'End Property

Public Property Let Enabled(ByVal NewVal As Boolean)
   ' No need to act unless this is a new value.
   If m_Enabled <> NewVal Then
      If m_hWnd Then
         ' This style can be toggled via the API.
         Call EnableWindow(m_hWnd, NewVal)
         m_Enabled = CBool(IsWindowEnabled(m_hWnd))
      Else
         m_Enabled = NewVal
      End If
      ' Attempt keeping buddy in sync.
      If m_hBuddy Then
         Call EnableWindow(m_hBuddy, m_Enabled)
      End If
   End If
End Property

Public Property Get Enabled() As Boolean
   ' Retrieve actual value.
   m_Enabled = CBool(IsWindowEnabled(m_hWnd))
   ' Attempt keeping buddy in sync.
   If m_hBuddy Then
      Call EnableWindow(m_hBuddy, m_Enabled)
   End If
   ' Return actual value.
   Enabled = m_Enabled
End Property

Public Property Let HotTrack(ByVal NewVal As Boolean)
   Dim NewStyle As Long
   
   ' Interpret style bit to use for this setting.
   If NewVal = True Then
      NewStyle = UDS_HOTTRACK
   Else
      NewStyle = 0&
   End If
   
   ' No need to cache/recreate unless this is a new value.
   If m_HotTrack <> NewStyle Then
      m_HotTrack = NewStyle
      ' This style can only be set when control is created.
      If m_hWnd Then Recreate
   End If
End Property

Public Property Get HotTrack() As Boolean
   ' Return cached value.
   HotTrack = CBool(m_HotTrack)
End Property

' TODO: Update to use full 32-bits if IE5+ is installed.
Public Property Let Max(ByVal NewVal As Long)
   ' Clamp to 16-bit range, for now.  Alternatives include
   ' testing against UD_MINVAL and UD_MAXVAL, but that's
   ' unsatisfactory because Min>Max is legal. So, just use
   ' loword and hope?  Nah...
   If NewVal > UD_MAXVAL Then
      NewVal = UD_MAXVAL
   ElseIf NewVal < UD_MINVAL Then
      NewVal = UD_MINVAL
   End If
   ' Update window if it exists, else just cache.
   If NewVal <> m_Max Then
      m_Max = NewVal
      Call RangeSet
   End If
End Property

Public Property Get Max() As Long
   ' Refresh cached values, and report.
   Call RangeGet
   Max = m_Max
End Property

' TODO: Update to use full 32-bits if IE5+ is installed.
Public Property Let Min(ByVal NewVal As Long)
   ' Clamp to 16-bit range, for now.  Alternatives include
   ' testing against UD_MINVAL and UD_MAXVAL, but that's
   ' unsatisfactory because Min>Max is legal. So, just use
   ' loword and hope?  Nah...
   If NewVal > UD_MAXVAL Then
      NewVal = UD_MAXVAL
   ElseIf NewVal < UD_MINVAL Then
      NewVal = UD_MINVAL
   End If
   ' Update window if it exists, else just cache.
   If NewVal <> m_Min Then
      m_Min = NewVal
      Call RangeSet
   End If
End Property

Public Property Get Min() As Long
   ' Refresh cached values, and report.
   Call RangeGet
   Min = m_Min
End Property

'Public Property Let Thousands(ByVal NewVal As Boolean)
'   Dim NewStyle As Long
'
'   ' Interpret style bit to use for this setting.
'   If NewVal = True Then
'      NewStyle = 0&
'   Else
'      NewStyle = UDS_NOTHOUSANDS
'   End If
'
'   ' No need to cache/recreate unless this is a new value.
'   If m_Thousands <> NewStyle Then
'      m_Thousands = NewStyle
'      ' This style can only be set when control is created.
'      If m_hWnd Then Recreate
'   End If
'End Property
'
'Public Property Get Thousands() As Boolean
'   ' Return inverse of cached value.
'   Thousands = Not CBool(m_Thousands)
'End Property

' TODO: Update to use UDM_SETPOS32 if IE5+ is installed.
Public Property Let Value(ByVal NewVal As Long)
   ' Although we accept a Long, we need to clamp it down
   ' to an Integer, by only using the LoWord.
   m_Value = LoWord(NewVal)
   ' Attempt to set new value, if window exists.
   If m_hWnd Then
      Call SendMessage(m_hWnd, UDM_SETPOS, 0&, ByVal MakeLong(m_Value, 0))
      ' Update cached value with new value.
      m_Value = Me.Value
   End If
End Property

' TODO: Update to use UDM_GETPOS32 if IE5+ is installed.
Public Property Get Value() As Long
   Dim nRet As Long
   ' Query actual control for current value if possible.
   If m_hWnd Then
      nRet = SendMessage(m_hWnd, UDM_GETPOS, 0&, ByVal 0&)
      ' High-order word contains error code.
      If HiWord(nRet) = 0 Then
         m_Value = LoWord(nRet)
      End If
   End If
   Value = m_Value
End Property

'Public Property Let Visible(ByVal NewVal As Boolean)
'   ' No need to do anything unless this is a new value.
'   If m_Visible <> NewVal Then
'      m_Visible = NewVal
'      ' Update actual visibility, if window exists.
'      If m_hWnd Then
'         If m_Visible Then
'            Call ShowWindow(m_hWnd, SW_SHOWNA)
'         Else
'            Call ShowWindow(m_hWnd, SW_HIDE)
'         End If
'         ' Confirm results.
'         m_Visible = CBool(IsWindowVisible(m_hWnd))
'      End If
'   End If
'End Property
'
'Public Property Get Visible() As Boolean
'   ' If window exists, check and cache actual value,
'   ' or else just return flag.
'   If m_hWnd Then
'      m_Visible = CBool(IsWindowVisible(m_hWnd))
'   End If
'   Visible = m_Visible
'End Property

'Public Property Let Wrap(ByVal NewVal As Boolean)
'   Dim NewStyle As Long
'
'   ' Interpret style bit to use for this setting.
'   If NewVal = True Then
'      NewStyle = UDS_WRAP
'   Else
'      NewStyle = 0&
'   End If
'
'   ' No need to cache/recreate unless this is a new value.
'   If m_Wrap <> NewStyle Then
'      m_Wrap = NewStyle
'      ' This style can only be set when control is created.
'      If m_hWnd Then Recreate
'   End If
'End Property
'
'Public Property Get Wrap() As Boolean
'   ' Return cached value.
'   Wrap = CBool(m_Wrap)
'End Property

' ***********************************************
'  Public Methods
' ***********************************************
Public Sub Refresh()
   Dim NewText As String
   ' Tell buddy to get with the program!
   NewText = Me.ToString
   Call SendMessage(m_hBuddy, WM_SETTEXT, 0&, ByVal NewText)
End Sub

Public Function ToString() As String
   ' Update current value, if available.
   m_Value = Me.Value
   
   ' Return a string representing the Value
   ' as it should appear in buddy control.
   Select Case m_Base
      Case udBase10
         If m_Thousands = UDS_NOTHOUSANDS Then
            ToString = Format$(m_Value, "0")
         Else
            ToString = Format$(m_Value, "#,##0")
         End If
      Case udbase16
         ' TODO: Revisit if/when 32-bit support is added.
         ToString = Hex$(m_Value)
   End Select
End Function

Public Sub Validate()
   Dim nRet As Long
   ' Validate user entry. UDM_GETPOS reports
   ' an error if the value is out of range or
   ' otherwise invalid (ex: alpha chars).
   nRet = SendMessage(m_hWnd, UDM_GETPOS, 0&, ByVal 0&)
   ' High-order word contains error code.
   If HiWord(nRet) <> 0 Then
      ' We need to restore valid value!
      If LenB(m_Buddy.Text) <> 0 Then
        Me.Value = m_Value
      End If
   Else
      ' Just cache current value, as this won't
      ' have updated if the user manually entered
      ' new data via keyboard or clipboard.
      m_Value = LoWord(nRet)
   End If
End Sub

' ***********************************************
'  Private Methods
' ***********************************************
Private Function AccelerationSet() As Boolean
   Dim ud(0 To 2) As UDACCEL
   Dim Elements As Long
   
   ' Bail if we don't have a window.
   If m_hWnd = 0 Then Exit Function
   
   ' In all cases, we want to start incrementing by one.
   ud(0).nInc = 1
   ud(0).nSec = 0
   
   ' Set up acceleration factors
   Select Case m_Acceleration
      Case udAccelerateNone
         ' Just the single element here.
         Elements = 1
         
      Case udAccelerateSlow
         ' Increment by 3 after 2 seconds
         Elements = 2
         ud(1).nInc = 3
         ud(1).nSec = 2
         
      Case udAccelerateMedium
         ' Increment by 3 after 2 seconds
         ' Increment by 5 after 4 seconds
         Elements = 3
         ud(1).nInc = 3
         ud(1).nSec = 2
         ud(2).nInc = 5
         ud(2).nSec = 4
         
      Case udAccelerateFast
         ' Increment by 10 after 2 seconds
         ' Increment by 25 after 4 seconds
         Elements = 3
         ud(1).nInc = 10
         ud(1).nSec = 2
         ud(2).nInc = 25
         ud(2).nSec = 4
   End Select
   
   ' Send the acceleration data to the control.
   AccelerationSet = CBool(SendMessage(m_hWnd, UDM_SETACCEL, ByVal Elements, ud(0)))
End Function

'Private Function FlipBit(ByVal Bit As Long, ByVal Value As Boolean) As Boolean
'   Dim nStyle As Long
'
'   ' Retrieve current style bits.
'   nStyle = GetWindowLong(m_hWnd, GWL_STYLE)
'
'   ' Attempt to set requested bit On or Off,
'   ' and redraw
'   If Value Then
'      nStyle = nStyle Or Bit
'   Else
'      nStyle = nStyle And Not Bit
'   End If
'   Call SetWindowLong(m_hWnd, GWL_STYLE, nStyle)
'   Call Redraw
'
'   ' Return success code.
'   FlipBit = (nStyle = GetWindowLong(m_hWnd, GWL_STYLE))
'End Function

Private Function HiWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function

Private Function LoWord(ByVal LongIn As Long) As Integer
   Call CopyMemory(LoWord, LongIn, 2)
End Function

Private Function MakeLong(ByVal WordLo As Integer, ByVal WordHi As Integer) As Long
   Call CopyMemory(MakeLong, WordLo, 2)
   Call CopyMemory(ByVal (VarPtr(MakeLong) + 2), WordHi, 2)
End Function

' TODO: Update to use UDM_GETRANGE32 if IE5+ is installed.
Private Sub RangeGet()
   Dim nRet As Long
   ' No need to bother unless the window exists.
   If m_hWnd Then
      ' Update cached range values.
      nRet = SendMessage(m_hWnd, UDM_GETRANGE, 0&, ByVal 0&)
      m_Max = LoWord(nRet)
      m_Min = HiWord(nRet)
   End If
End Sub

' TODO: Update to use UDM_SETRANGE32 if IE5+ is installed.
Private Sub RangeSet()
   ' No need to bother unless the window exists.
   If m_hWnd Then
      ' Attempt setting new range, and check results.
      Call SendMessage(m_hWnd, UDM_SETRANGE, 0&, ByVal MakeLong(m_Max, m_Min))
      Call RangeGet
      ' Changing the range may also change the position; update.
      m_Value = Me.Value
   End If
End Sub

Private Function Recreate() As Boolean
   Dim nStyle As Long
   Dim pt As POINTAPI
   
   ' No need to continue if we don't have a buddy.
   If IsWindow(m_hBuddy) = 0 Then Exit Function
   
   ' Some housekeeping required, if window already exists.
   If IsWindow(m_hWnd) Then
      ' Grab current values, so they can be restored on recreate.
      Call RangeGet
      m_Value = Me.Value
      ' Kill existing spinner.
      Call DestroyWindow(m_hWnd)
      ' Reset handle to zero.
      m_hWnd = 0
      ' Restore original position of buddy textbox.
      With m_OrigRect
         Call MoveWindow(m_hBuddy, .Left, .Top, .Right, .Bottom, True)
      End With
   End If
   
   ' Cache size/position of buddy so it can be restored later.
   Call GetWindowRect(m_hBuddy, m_OrigRect)
   With m_OrigRect
      ' Store width in .Right and height in .Bottom
      .Right = .Right - .Left
      .Bottom = .Bottom - .Top
      ' Convert screen coords to client coords.
      pt.X = .Left
      pt.Y = .Top
      Call ScreenToClient(GetParent(m_hBuddy), pt)
      .Left = pt.X
      .Top = pt.Y
   End With
   
   ' The purpose of this class is to buddy-up an up/down
   ' with a textbox, so this style is a given.
   nStyle = UDS_SETBUDDYINT
   
   ' Standard window styles.
   nStyle = nStyle Or WS_CHILD
   
   ' Style flags directly mapped to member variables.
   nStyle = nStyle Or m_Align       ' UDS_ALIGNLEFT / UDS_ALIGNRIGHT
   nStyle = nStyle Or m_ArrowKeys   ' UDS_ARROWKEYS
   nStyle = nStyle Or m_Direction   ' UDS_HORZ / 0&
   nStyle = nStyle Or m_HotTrack    ' UDS_HOTTRACK / 0&
   nStyle = nStyle Or m_Thousands   ' UDS_NOTHOUSANDS / 0&
   nStyle = nStyle Or m_Wrap        ' UDS_WRAP / 0&
   
   ' Style flags mapped to Boolean variables.
   If m_Border Then nStyle = nStyle Or WS_BORDER
   If Not m_Enabled Then nStyle = nStyle Or WS_DISABLED
   If m_Visible Then nStyle = nStyle Or WS_VISIBLE
   
   ' Create control.
   m_hWnd = CreateUpDownControl(nStyle, 0&, 0&, 0&, 0&, GetParent(m_hBuddy), 0&, App.hInstance, m_hBuddy, m_Max, m_Min, m_Value)
   
   ' Set appropriate radix base.
   If m_hWnd Then
      Call SendMessage(m_hWnd, UDM_SETBASE, m_Base, ByVal 0&)
   End If
   
   ' Set the appropriate acceleration rate.
   Call AccelerationSet
   
   ' Make sure buddy's Enabled property matches ours.
   Call EnableWindow(m_hBuddy, Enabled)
   
   ' Return results
   Recreate = CBool(m_hWnd)
End Function

'Private Sub Redraw()
'   ' Redraw window with new style.
'   Const swpFlags As Long = _
'      SWP_FRAMECHANGED Or SWP_NOMOVE Or _
'      SWP_NOZORDER Or SWP_NOSIZE
'   SetWindowPos m_hWnd, 0, 0, 0, 0, 0, swpFlags
'End Sub

'Private Function PointerToDWord(ByVal lpDWord As Long) As Long
'   Dim nRet As Long
'   If lpDWord Then
'      CopyMemory nRet, ByVal lpDWord, 4
'      PointerToDWord = nRet
'   End If
'End Function

'Private Function Style(Optional ByVal NewBits As Long = 0) As Long
'   ' Attempt to set new style bits.
'   If NewBits Then
'      Call SetWindowLong(m_hWnd, GWL_STYLE, NewBits)
'   End If
'   ' Retrieve current style bits.
'   Style = GetWindowLong(m_hWnd, GWL_STYLE)
'End Function


