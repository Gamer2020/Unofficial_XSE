VERSION 5.00
Begin VB.UserControl TabControl 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   ControlContainer=   -1  'True
   ScaleHeight     =   495
   ScaleWidth      =   2685
   ToolboxBitmap   =   "cTabCtrl.ctx":0000
End
Attribute VB_Name = "TabControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ======================================================================
' Declares and types:
' ======================================================================
' Windows general:
'Private Const WM_DESTROY = &H2
'Private Const WM_SETFOCUS = &H7
'Private Const WM_PAINT = &HF
'Private Const WM_ERASEBKGND = &H14
'Private Const WM_MOUSEACTIVATE = &H21
'Private Const WM_DRAWITEM = &H2B
Private Const WM_NOTIFY = &H4E
'Private Const WM_NCPAINT = &H85
'Private Const WM_KEYDOWN = &H100
'Private Const WM_USER = &H400

'Private Const MA_NOACTIVATE = 3

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
'Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
'Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
'Private Declare Function GetFocus Lib "user32" () As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function InvalidateRectAsNull Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, lpRect As Any, ByVal bErase As Long) As Long
'Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const SW_HIDE = 0
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOZORDER = &H4
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
'Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
'Private Const WS_BORDER = &H800000
Private Const WM_SETFONT = &H30
Private Const GWL_STYLE = (-16)

' Font
Private Const LF_FACESIZE = 32

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
'Private Const FF_DONTCARE = 0
'Private Const DEFAULT_QUALITY = 0
'Private Const DEFAULT_PITCH = 0
'Private Const DEFAULT_CHARSET = 1

Private Declare Function CreateFontIndirect& Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT)
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
    
'Private Const BITSPIXEL = 12
'Private Const LOGPIXELSX = 88    '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90    '  Logical pixels/inch in Y

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Private Const ODS_SELECTED = &H1
'Private Const ODT_HEADER = 100
'Private Const ODT_TAB = 101
'Private Const ODT_LISTVIEW = 102

'Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
'Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
'Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
'Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
'Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
'Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long

'Private Const TRANSPARENT = 1

'Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
'Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
'Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'Private Const PS_DASH = 1
'Private Const PS_DASHDOT = 3
'Private Const PS_DASHDOTDOT = 4
'Private Const PS_DOT = 2
'Private Const PS_SOLID = 0
'Private Const PS_NULL = 5

'Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
'Private Const DT_CENTER = &H1
'Private Const DT_VCENTER = &H4
'Private Const DT_SINGLELINE = &H20

'Private Type PAINTSTRUCT
'   hDC As Long
'   fErase As Long
'   rcPaint As RECT
'   fRestore As Long
'   fIncUpdate As Long
'   rgbReserved(0 To 31) As Byte
'End Type

'Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
'Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long

Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Private Const TCM_FIRST = &H1300&                   '// Tab control messages
'Private Const CCM_FIRST = &H2000                   '// Common control shared messages
'Private Const CCM_SETUNICODEFORMAT = (CCM_FIRST + 5)
'Private Const CCM_GETUNICODEFORMAT = (CCM_FIRST + 6)
Private Const H_MAX As Long = &HFFFF + 1
Private Const TCN_FIRST = H_MAX - 550                  '// tab control
'Private Const NM_FIRST = H_MAX
'Private Const NM_RCLICK = (NM_FIRST - 5)               '// uses NMCLICK struct

'Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)         '// lParam is bkColor

'Private Type COLORSCHEME
'   dwSize As Long
'   clrBtnHighlight As Long '       // highlight color
'   clrBtnShadow As Long          '// shadow color
'End Type

'Private Const CCM_SETCOLORSCHEME = (CCM_FIRST + 2)     '// lParam is color scheme
'Private Const CCM_GETCOLORSCHEME = (CCM_FIRST + 3)     '// fills in COLORSCHEME pointed to by lParam

'Private Const TTN_FIRST = (H_MAX - 520&)
'Private Const TTN_NEEDTEXTA = (TTN_FIRST - 0&)
'Private Const TTN_NEEDTEXT = TTN_NEEDTEXTA
'Private Const TTM_ACTIVATE = (WM_USER + 1)

' //====== TAB CONTROL ==========================================================
Private Const WC_TABCONTROLA = "SysTabControl32"
Private Const WC_TABCONTROL = WC_TABCONTROLA

'Private Const TCS_SCROLLOPPOSITE = &H1          ' // assumes multiline tab
'Private Const TCS_BOTTOM = &H2
'Private Const TCS_RIGHT = &H2
'Private Const TCS_MULTISELECT = &H4            ' // allow multi-select in button mode
'Private Const TCS_FLATBUTTONS = &H8
'Private Const TCS_FORCELABELLEFT = &H20&
Private Const TCS_HOTTRACK = &H40&
'Private Const TCS_VERTICAL = &H80&
'Private Const TCS_TABS = &H0
Private Const TCS_BUTTONS = &H100&
Private Const TCS_SINGLELINE = &H0
Private Const TCS_MULTILINE = &H200&
'Private Const TCS_FIXEDWIDTH = &H400&
'Private Const TCS_RAGGEDRIGHT = &H800&
'Private Const TCS_FOCUSONBUTTONDOWN = &H1000&
Private Const TCS_FOCUSNEVER = &H8000&
'Private Const TCS_EX_REGISTERDROP = &H2

Private Const TCM_GETITEMCOUNT = (TCM_FIRST + 4)

Private Const TCIF_TEXT = &H1
Private Const TCIF_IMAGE = &H2
Private Const TCIF_RTLREADING = &H4
Private Const TCIF_PARAM = &H8
Private Const TCIF_STATE = &H10

Private Const TCIS_BUTTONPRESSED = &H1
Private Const TCIS_HIGHLIGHTED = &H2

'Private Type TCITEMHEADER
'    mask As Long
'    lpReserved1 As Long
'    lpReserved2 As Long
'    pszText As String
'    cchTextMax As Long
'    iImage As Long
'End Type

'Private Type TCITEMHEADER_NOTEXT
'    mask As Long
'    lpReserved1 As Long
'    lpReserved2 As Long
'    pszText As Long
'    cchTextMax As Long
'    iImage As Long
'End Type

Private Type TCITEM
    mask As Long
    dwState As Long
    dwStateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
End Type

Private Const TCM_GETITEMA = (TCM_FIRST + 5)
'Private Const TCM_GETITEMW = (TCM_FIRST + 60)
'Private Const TCM_GETITEM = TCM_GETITEMA
Private Const TCM_SETITEMA = (TCM_FIRST + 6)
'Private Const TCM_SETITEMW = (TCM_FIRST + 61)
'Private Const TCM_SETITEM = TCM_SETITEMA
Private Const TCM_INSERTITEMA = (TCM_FIRST + 7)
'Private Const TCM_INSERTITEMW = (TCM_FIRST + 62)
Private Const TCM_INSERTITEM = TCM_INSERTITEMA
Private Const TCM_DELETEITEM = (TCM_FIRST + 8)
Private Const TCM_DELETEALLITEMS = (TCM_FIRST + 9)
'Private Const TCM_GETITEMRECT = (TCM_FIRST + 10)
Private Const TCM_GETCURSEL = (TCM_FIRST + 11)
Private Const TCM_SETCURSEL = (TCM_FIRST + 12)
'Private Const TCHT_NOWHERE = &H1
'Private Const TCHT_ONITEMICON = &H2
'Private Const TCHT_ONITEMLABEL = &H4
'Private Const TCHT_ONITEM = (TCHT_ONITEMICON Or TCHT_ONITEMLABEL)

'Private Type TCHITTESTINFO
'    pt As POINTAPI
'    flags As Long
'End Type

'Private Const TCM_HITTEST = (TCM_FIRST + 13)
'Private Const TCM_SETITEMEXTRA = (TCM_FIRST + 14)
Private Const TCM_ADJUSTRECT = (TCM_FIRST + 40)
'Private Const TCM_SETITEMSIZE = (TCM_FIRST + 41)
'Private Const TCM_REMOVEIMAGE = (TCM_FIRST + 42)
Private Const TCM_SETPADDING = (TCM_FIRST + 43)
'Private Const TCM_GETROWCOUNT = (TCM_FIRST + 44)
'Private Const TCM_GETTOOLTIPS = (TCM_FIRST + 45)
'Private Const TCM_SETTOOLTIPS = (TCM_FIRST + 46)
'Private Const TCM_GETCURFOCUS = (TCM_FIRST + 47)
'Private Const TCM_SETCURFOCUS = (TCM_FIRST + 48)
'Private Const TCM_SETMINTABWIDTH = (TCM_FIRST + 49)
'Private Const TCM_DESELECTALL = (TCM_FIRST + 50)
'Private Const TCM_HIGHLIGHTITEM = (TCM_FIRST + 51)
'Private Const TCM_SETEXTENDEDSTYLE = (TCM_FIRST + 52)    ' // optional wParam == mask
'Private Const TCM_GETEXTENDEDSTYLE = (TCM_FIRST + 53)
'Private Const TCM_SETUNICODEFORMAT = CCM_SETUNICODEFORMAT
'Private Const TCM_GETUNICODEFORMAT = CCM_GETUNICODEFORMAT
Private Const TCN_KEYDOWN = (TCN_FIRST - 0)

Private Type TCKEYDOWN
    hdr As NMHDR
    B(0 To 5) As Byte
End Type

Private Const TCN_SELCHANGE = (TCN_FIRST - 1)
Private Const TCN_SELCHANGING = (TCN_FIRST - 2)
'Private Const TCN_GETOBJECT = (TCN_FIRST - 3)

' ======================================================================
' Interface:
' ======================================================================
Private m_hWnd As Long
Private m_hWndCtl As Long
'Private m_hIml As Long
Private m_sKey() As String
Private m_tULF As LOGFONT
Private m_hFnt As Long

Private m_HotTrack As Boolean
Private m_Buttons As Boolean
Private m_MultiLine As Boolean
Private m_FocusRect As Boolean

Private cSubclasser As cSelfSubclasser
'Private m_InIDE As Boolean

Public Event BeforeClick(ByVal lTab As Long, ByRef Cancel As Boolean)
Public Event TabClick(ByVal lTab As Long)
Attribute TabClick.VB_Description = "Raised when a tab is clicked."

Public Property Get HotTrack() As Boolean
Attribute HotTrack.VB_Description = "Gets/sets whether tab control tracks the mouse and highlights tabs pointed to by the cursor or not. If set at run-time, call the Rebuild method to recreate the control with the new style."
   HotTrack = m_HotTrack
End Property

Public Property Let HotTrack(ByVal bState As Boolean)
   m_HotTrack = bState
   pChangeStyle
   PropertyChanged "HotTrack"
End Property

Public Property Get Buttons() As Boolean
Attribute Buttons.VB_Description = "Gets/sets whether the tabs appear as buttons instead of tabs. If set at run-time, call the Rebuild method to recreate the control with the new style."
   Buttons = m_Buttons
End Property

Public Property Let Buttons(ByVal bState As Boolean)
   m_Buttons = bState
   pChangeStyle
   PropertyChanged "Buttons"
End Property

Public Property Get MultiLine() As Boolean
Attribute MultiLine.VB_Description = "Gets/sets whether tabs appear on more than one line or not. If changed at run-time, call the Rebuild method to recreate the control with the new style."
   MultiLine = m_MultiLine
End Property

Public Property Let MultiLine(ByVal bState As Boolean)
   m_MultiLine = bState
   pChangeStyle
   PropertyChanged "MultiLine"
End Property

Public Property Get FocusRect() As Boolean
   FocusRect = m_FocusRect
End Property

Public Property Let FocusRect(ByVal bState As Boolean)
   m_FocusRect = bState
   pChangeStyle
   PropertyChanged "FocusRect"
End Property

Public Sub SetPadding(ByVal xPixels As Long, ByVal yPixels As Long)
Dim lXY As Long
   lXY = xPixels Or ((yPixels And &H7FFF) * &H10000)
   SendMessageLong m_hWnd, TCM_SETPADDING, 0, lXY
End Sub

Public Property Get font() As StdFont
Attribute font.VB_Description = "Gets/sets the font used by the tab control."
    Set font = UserControl.font
End Property

Public Property Set font(sFont As StdFont)
   If Not (UserControl.font Is sFont) Then
      Set UserControl.font = sFont
      pSetFont sFont
      PropertyChanged "Font"
   End If
End Property

Private Sub pSetFont(ByRef sFont As StdFont)
Dim hFnt As Long
   
   ' Store a log font structure for this font:
   pOLEFontToLogFont sFont, UserControl.hDC, m_tULF
   
   ' Store old font handle:
   hFnt = m_hFnt
   
   ' Create a new version of the font:
   m_hFnt = CreateFontIndirect(m_tULF)
   
   ' Ensure the edit portion has the correct font:
   If (m_hWnd <> 0) Then
       SendMessage m_hWnd, WM_SETFONT, m_hFnt, 1
   End If
   
   ' Delete previous version, if we had one:
   If (hFnt <> 0) Then
       DeleteObject hFnt
   End If
   
End Sub

Private Sub pOLEFontToLogFont(fntThis As StdFont, hDC As Long, tLF As LOGFONT)
Dim sFont As String
Dim iChar As Integer

    ' Convert an OLE StdFont to a LOGFONT structure:
    With tLF
        
        sFont = fntThis.name
        
        ' There is a quicker way involving StrConv and CopyMemory, but
        ' this is simpler!:
        For iChar = 1 To Len(sFont)
            .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
        Next iChar
        
        ' Based on the Win32SDK documentation:
        .lfHeight = -MulDiv((fntThis.SIZE), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
        .lfItalic = fntThis.Italic
        
        If (fntThis.Bold) Then
            .lfWeight = FW_BOLD
        Else
            .lfWeight = FW_NORMAL
        End If
        
        .lfUnderline = fntThis.Underline
        .lfStrikeOut = fntThis.Strikethrough
        
    End With

End Sub

Public Sub AddTab( _
        ByVal sText As String, _
        Optional ByVal vKeyBefore As Variant = -1, _
        Optional ByVal sKey As String)
Attribute AddTab.VB_Description = "Adds or inserts a tab."
Dim tTCI As TCITEM
Dim lTabCount As Long
Dim lKey As Long
Dim lIndex As Long
Dim lRet As Long

   ' Set up the tab to add
   lTabCount = TabCount
   
    With tTCI
      .lParam = lTabCount
      .mask = TCIF_TEXT
      .cchTextMax = Len(sText) + 1
      .pszText = sText & " "
    End With
    
    ReDim Preserve m_sKey(0 To lTabCount) As String
   
   If Not (IsNumeric(vKeyBefore)) Then
      lIndex = APITabIndex(vKeyBefore)
   ElseIf (vKeyBefore > -1) Then
      lIndex = vKeyBefore - 1
   Else
      lIndex = lTabCount
   End If
        
   ' Add the tab
   lRet = SendMessage(m_hWnd, TCM_INSERTITEM, lIndex, tTCI)
   SelectTab lTabCount + 1
   
   If (lRet <> lIndex) <> 0 Then
   
       ' Add the key
       For lKey = lTabCount To lIndex + 1 Step -1
           m_sKey(lKey) = m_sKey(lKey - 1)
       Next lKey
       
       m_sKey(lIndex) = sKey
       
       
       
   End If

End Sub

Public Sub RemoveTab(ByVal vKey As Variant)
Attribute RemoveTab.VB_Description = "Removes a tab from the control."
Dim lIndex As Long
Dim lR As Long
Dim i As Long
Dim bSelected As Boolean

   lIndex = APITabIndex(vKey)
   bSelected = (SelectedTab - 1 = lIndex)
   lR = SendMessageLong(m_hWnd, TCM_DELETEITEM, lIndex, 0)

   If (lR <> 0) Then
      
      If TabCount > 0 Then
         
         If (bSelected) Then
         
           If (lIndex - 1 > 0) Then
              SelectTab lIndex
           Else
              SelectTab 1
           End If
           
        End If
         
         For i = lIndex To UBound(m_sKey) - 1
            m_sKey(i) = m_sKey(i + 1)
         Next i
         
         ReDim Preserve m_sKey(0 To TabCount - 1) As String

      Else
         Erase m_sKey
      End If
      
   End If
   
End Sub

Public Sub RemoveAllTabs()
Attribute RemoveAllTabs.VB_Description = "Removes all tabs from the control."

   SendMessageLong m_hWnd, TCM_DELETEALLITEMS, 0, 0
   
   If (TabCount = 0) Then
      Erase m_sKey
   End If
   
End Sub

Public Property Get SelectedTab() As Long
Attribute SelectedTab.VB_Description = "Gets the index of the selected tab."
    SelectedTab = SendMessageLong(m_hWnd, TCM_GETCURSEL, 0, 0) + 1
End Property

Public Sub SelectTab(ByVal vKey As Variant, Optional ByVal NoEvents As Boolean = False)
Attribute SelectTab.VB_Description = "Selects a tab in the control."
Dim lR As Long
Dim Cancel As Boolean
Dim lIndex As Long

   lIndex = APITabIndex(vKey)
   
   If (lIndex > -1) Then
   
      If (Not (NoEvents)) Then
      
         If (SelectedTab > 0) Then
            RaiseEvent BeforeClick(SelectedTab, Cancel)
         End If
         
      End If
      
      If Not (Cancel) Then
      
         lR = SendMessageLong(m_hWnd, TCM_SETCURSEL, lIndex, 0)
         
         If (lR >= -1) Then
         
            If Not (NoEvents) Then
                RaiseEvent TabClick(lIndex + 1)
            End If
            
         End If
         
      End If
      
   End If
   
End Sub

Public Sub Rebuild()
Attribute Rebuild.VB_Description = "Rebuilds the tab control.  Use this if you change any of the style properties at run-time to allow the style change to take effect."
Dim i As Long
Dim tTI() As TCITEM
Dim iICount As Long
Dim tR As RECT
Dim lTab As Long

   iICount = TabCount
   If (iICount > 0) Then
      ReDim tTI(0 To iICount - 1) As TCITEM
      For i = 0 To iICount - 1
         With tTI(i)
            .mask = TCIF_IMAGE Or TCIF_TEXT Or TCIF_PARAM Or TCIF_STATE Or TCIF_RTLREADING
            .cchTextMax = 255
            .pszText = String$(255, 0)
            .dwStateMask = TCIS_BUTTONPRESSED
         End With
         SendMessage m_hWnd, TCM_GETITEMA, i, tTI(i)
      Next i

      lTab = SelectedTab
      
   End If
   
   pTerminate
   pInitialise
   
   If (iICount > 0) Then
      For i = 0 To iICount - 1
         SendMessage m_hWnd, TCM_INSERTITEM, i + 1, tTI(i)
      Next i
   End If
   
   pSetFont UserControl.font
   GetWindowRect m_hWnd, tR
   SetWindowPos m_hWnd, 0, 0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, SWP_NOMOVE Or SWP_FRAMECHANGED Or SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_NOOWNERZORDER

   SelectTab lTab, True
      
End Sub

Public Property Get ClientLeft() As Long
Attribute ClientLeft.VB_Description = "Gets the left position of the client area of the tab control."
Dim RC As RECT
    pGetClientRect RC
    ClientLeft = RC.Left * Screen.TwipsPerPixelX
End Property

Public Property Get ClientTop() As Long
Attribute ClientTop.VB_Description = "Gets the top position of the client area of the tab control."
Dim RC As RECT
    pGetClientRect RC
    ClientTop = RC.Top * Screen.TwipsPerPixelY
End Property

Public Property Get ClientWidth() As Long
Attribute ClientWidth.VB_Description = "Gets the width of the client area of the tab control."
Dim RC As RECT
    pGetClientRect RC
    ClientWidth = (RC.Right - RC.Left) * Screen.TwipsPerPixelX
End Property

Public Property Get ClientHeight() As Long
Attribute ClientHeight.VB_Description = "Gets the height of the client area of the tab control."
Dim RC As RECT
    pGetClientRect RC
    ClientHeight = (RC.Bottom - RC.Top) * Screen.TwipsPerPixelY
End Property

Private Sub pGetClientRect(RC As RECT)
Dim tP As POINTAPI
    
    ' Get window rect of the user control:
    GetWindowRect m_hWndCtl, RC
    tP.x = RC.Left
    tP.Y = RC.Top
    
    ' Adjust to coordinates of user control's container:
    ScreenToClient GetParent(m_hWndCtl), tP
    RC.Right = RC.Right + (tP.x - RC.Left)
    RC.Bottom = RC.Bottom + (tP.Y - RC.Top)
    RC.Left = tP.x
    RC.Top = tP.Y
    
    ' Calculate the useable area of the tab:
    SendMessage m_hWnd, TCM_ADJUSTRECT, 0, RC
    
End Sub

Public Property Get TabText(ByVal vKey As Variant) As String
Attribute TabText.VB_Description = "Gets/sets the text which appears in a tab."
Dim lIndex As Long
Dim tTI As TCITEM
Dim lR As Long
Dim sText As String

    lIndex = APITabIndex(vKey)
    tTI.cchTextMax = 255
    tTI.pszText = String$(255, 0)
    tTI.mask = TCIF_TEXT
    lR = SendMessage(m_hWnd, TCM_GETITEMA, lIndex, tTI)
    
    If (lR <> 0) Then
       
       sText = tTI.pszText
       lR = InStrB(sText, vbNullChar)
       
       If (lR <> 0) Then
          TabText = RTrim$(LeftB$(sText, lR - 1))
       Else
          TabText = sText
       End If
       
    End If
    
End Property

Public Property Let TabText(ByVal vKey As Variant, ByVal sText As String)
Dim lIndex As Long
Dim tTI As TCITEM

    lIndex = APITabIndex(vKey)
    tTI.cchTextMax = Len(sText)
    tTI.pszText = sText & " "
    tTI.mask = TCIF_TEXT
    SendMessage m_hWnd, TCM_SETITEMA, lIndex, tTI
    
End Property

Public Property Get TabHot(ByVal vKey As Variant) As Boolean
Dim lIndex As Long
Dim tTI As TCITEM
Dim lR As Long
   
    lIndex = APITabIndex(vKey)
    
    If (lIndex > -1) Then
    
        tTI.mask = TCIF_STATE
        lR = SendMessage(m_hWnd, TCM_GETITEMA, lIndex, tTI)
        
        If lR <> 0 Then
            TabHot = ((tTI.dwState And TCIS_HIGHLIGHTED) = TCIS_HIGHLIGHTED)
        End If
    
    End If
   
End Property

Public Property Get TabKey(ByVal lIndex As Long)
Attribute TabKey.VB_Description = "Gets/sets the key to associate with a tab."
    If (lIndex > 0) And (lIndex <= TabCount) Then
        TabKey = m_sKey(lIndex - 1)
    End If
End Property

Private Property Get APITabIndex(ByVal vKey As Variant) As Long
    APITabIndex = IndexForTab(vKey) - 1
End Property

Public Property Get IndexForTab(ByVal vKey As Variant) As Long
Attribute IndexForTab.VB_Description = "Gets the numeric index of a tab given the key."
Dim lS As Long
Dim lKey As Long
    
    lKey = -1
    
    If IsNumeric(vKey) Then
        lKey = CLng(vKey) - 1
    Else
        For lS = 0 To TabCount - 1
            If (m_sKey(lS) = vKey) Then
                lKey = lS
                Exit For
            End If
        Next lS
    End If
    
    If (lKey >= 0) And (lKey < TabCount) Then
        IndexForTab = lKey + 1
    End If

End Property

Public Property Get TabCount() As Long
Attribute TabCount.VB_Description = "Gets the number of tabs in the control."
    TabCount = SendMessageLong(m_hWnd, TCM_GETITEMCOUNT, 0, 0)
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Gets the Window handle of the control.  Use TabCtrlhWnd if you want the hWnd of the tab itself."
    hWnd = m_hWndCtl
End Property

Public Property Get TabCtrlhWnd() As Long
Attribute TabCtrlhWnd.VB_Description = "Gets the hWnd of the Tab Control."
   TabCtrlhWnd = m_hWnd
End Property

Private Sub pChangeStyle()
Dim dwStyle As Long
   
    If m_hWnd <> 0 Then
        
        If (m_HotTrack) Then
            dwStyle = TCS_HOTTRACK
        End If
      
        If (m_MultiLine) Then
            dwStyle = dwStyle Or TCS_MULTILINE
        Else
            dwStyle = dwStyle Or TCS_SINGLELINE
        End If
        
        If Not (m_FocusRect) Then
            dwStyle = dwStyle Or TCS_FOCUSNEVER
        End If
       
       ' Create the control:
       dwStyle = dwStyle Or WS_VISIBLE Or WS_CHILD Or WS_CLIPSIBLINGS
      
      SetWindowLong m_hWnd, GWL_STYLE, dwStyle
      
   End If
End Sub

Private Sub pInitialise()
Dim dwStyle As Long
Dim tR As RECT
        
    ' Ensure we don't already have Tab control
    pTerminate
    
    If (m_HotTrack) Then
        dwStyle = TCS_HOTTRACK
    End If
   
    If (m_Buttons) Then
        dwStyle = dwStyle Or TCS_BUTTONS
    End If
   
    If (m_MultiLine) Then
        dwStyle = dwStyle Or TCS_MULTILINE
    Else
        dwStyle = dwStyle Or TCS_SINGLELINE
    End If
   
    If Not (m_FocusRect) Then
        dwStyle = dwStyle Or TCS_FOCUSNEVER
    End If
    
    ' Create the control
    dwStyle = dwStyle Or WS_VISIBLE Or WS_CHILD Or WS_CLIPSIBLINGS
    
    m_hWndCtl = UserControl.hWnd
    GetClientRect m_hWndCtl, tR
    m_hWnd = CreateWindowEx( _
        0, WC_TABCONTROL, "", _
        dwStyle, _
        0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, _
        m_hWndCtl, 0, _
        App.hInstance, 0)
        
    If (m_hWnd <> 0) Then
        
        If (UserControl.Ambient.UserMode) Then
        
            Set cSubclasser = New cSelfSubclasser
        
            If cSubclasser.ssc_Subclass(m_hWndCtl, , 1, Me) = True Then
                cSubclasser.ssc_AddMsg m_hWndCtl, eMsgWhen.MSG_AFTER, eAllMessages.ALL_MESSAGES
            End If
        
        End If
        
        AddTab "Script1"
        
    End If
    
End Sub

Private Sub pTerminate()
   
    If (m_hWnd <> 0) Then
        
        If (UserControl.Ambient.UserMode) Then
            ' Stop subclassing
            Set cSubclasser = Nothing
        End If
        
        ' Destroy the window
        ShowWindow m_hWnd, SW_HIDE
        SetParent m_hWnd, 0
        DestroyWindow m_hWnd
        
        ' store that we haven't a window
        m_hWnd = 0
        
    End If
    
   ' Clear up font:
   If (m_hFnt <> 0) Then
        DeleteObject m_hFnt
        m_hFnt = 0
   End If
   
End Sub

Private Sub UserControl_InitProperties()
    'pInitialise
    Set font = UserControl.Ambient.font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    m_HotTrack = PropBag.ReadProperty("HotTrack", False)
    m_Buttons = PropBag.ReadProperty("Buttons", False)
    m_MultiLine = PropBag.ReadProperty("MultiLine", False)
    
    pInitialise
    
    Dim sFnt As New StdFont
    
    sFnt.name = "Tahoma"
    sFnt.SIZE = 8
    
    Set font = PropBag.ReadProperty("Font", sFnt)
    
End Sub

Private Sub UserControl_Resize()
Dim tR As RECT
   If (m_hWnd <> 0) Then
      GetClientRect m_hWndCtl, tR
      MoveWindow m_hWnd, 0, 0, tR.Right - tR.Left, tR.Bottom - tR.Top, 1
   End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    'pTerminate
    
    Dim sFnt As New StdFont
    sFnt.name = "Tahoma"
    sFnt.SIZE = 8
    
    PropBag.WriteProperty "Font", font, sFnt
    PropBag.WriteProperty "HotTrack", m_HotTrack, False
    PropBag.WriteProperty "Buttons", m_Buttons, False
    PropBag.WriteProperty "MultiLine", m_MultiLine, False
    
End Sub

'- ordinal #1
Private Sub myWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc.
'*              Not applicable with After messages
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
 
Dim tNM As NMHDR
Dim tTKD As TCKEYDOWN
Dim wKey As Long
Dim lTab As Long
Dim lFlags As Long
Dim Cancel As Boolean
    
    Select Case uMsg
                
        Case WM_NOTIFY
                    
        CopyMemory tNM, ByVal lParam, Len(tNM)
        
        If (tNM.hwndFrom = m_hWnd) Then
        
            Select Case tNM.code
            
                Case TCN_KEYDOWN
                    
                    CopyMemory tTKD, ByVal lParam, Len(tTKD)
                    wKey = tTKD.B(1) * &H100& Or tTKD.B(0)
                    CopyMemory lFlags, tTKD.B(2), 4
        
                Case TCN_SELCHANGING
                    
                    lTab = SelectedTab
                    
                    If (lTab <> 0) Then
                       
                       RaiseEvent BeforeClick(lTab, Cancel)
                       
                       If (Cancel) Then
                          bHandled = True
                          lReturn = 1
                       End If
                       
                    End If
                    
                Case TCN_SELCHANGE
                
                    lTab = SelectedTab
                    RaiseEvent TabClick(lTab)
                    
            End Select
            
        End If
            
    End Select

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
' *************************************************************
        
End Sub
