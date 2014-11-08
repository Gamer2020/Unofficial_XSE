VERSION 5.00
Begin VB.UserControl JCToolbar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8235
   ControlContainer=   -1  'True
   PropertyPages   =   "JCToolbar.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   549
   ToolboxBitmap   =   "JCToolbar.ctx":0013
   Begin VB.Timer tmrTooltip 
      Enabled         =   0   'False
      Interval        =   4900
      Left            =   6840
      Top             =   0
   End
   Begin VB.PictureBox PicRight 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   8595
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   0
      Top             =   0
      Width           =   195
      Begin VB.Timer tmrBtns 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   0
         Top             =   0
      End
   End
   Begin VB.PictureBox PicTB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   135
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   527
      TabIndex        =   2
      Top             =   0
      Width           =   7905
   End
   Begin VB.PictureBox PicLeft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   1
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "JCToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'=========================================================================
'
'   jcToolbars v 2.0.1
'   Copyright © 2005 Juan Carlos San Román Arias (sanroman2004@yahoo.com)
'
'=========================================================================

'   ------------------------------
'   Version 2.0 Data: 25-Dec-2005
'   ------------------------------
'   - It is just one control (it includes toolbar button (improved JCF_Toolbutton created by João Fortes) and vertical 3d line)
'   - add buttons and separators at runtime
'   - popup menu that shows hidden buttons and five theme colors
'   - auto hide toolbar buttons when toolbar is sizing
'   - initial toolbar autosizing (height and width) taking into account
'     icon and font size used in toolbar buttons.
'   - added two customs theme colors (norton 2004 and visual studio 2005 themes)
'   - It´s possible to change at runtime:
'     - button caption
'     - button state
'     - button tag
'     - button tooltiptext
'     - button value
'     - menu language selection (english and spanish)
'     - show or hide menu to change toolbar theme color
'   - improving of JCF_Toolbutton created by João Fortes
'     - icon size can be changed
'     - font can be changed (type, size, bold, italic)
'     - font color can be changed
'     - added four type of icon and caption aligments  color can be changed (IconLeftTextRight, IconRightTextLeft, IconTopTextBotton and IconBottonTextTop)

'   ------------------------------
'   Version 1.0  Data: 23-Nov-2005
'   ------------------------------
'   This is an Office 2003 toolbar for VB. You can built a nice Toolbar.
'   The initial idea taken from JCF_Toolbutton created by João Fortes.
'   I have made a compilation of different jobs published on Planet-Source-Code.com
'   I want to thank to
'   - Everyday Panos for your Office 2003 Button AND MOVING TOOLBAR project
'   - Fred cpp for api functions used in his isbutton control
'   - Carles P.V. for 3d UcVertical line
'   - All control is drawn using api functions (no images, no other controls)
'
'   ---------------------------------
'   Version 2.0.1  Data: 15-Feb-2006
'   ---------------------------------
'   - Control sucture was completely reorganized (just one control)
'   - Chevrons have been added to show hidden buttons and 5 theme colors
'   - Buttons and separators can be added, moved or deleted at design and runtime
'   - Toolbar autosizing taking into account icon and font size used in toolbar buttons
'   - Windows XP theme auto detection or selection  (blue, silver and olive)
'   - Two customs theme colors (norton 2004 and visual studio 2005 themes) have been added
'   - You can select language for theme color Menu (spanish and english)
'   - You can determine if theme color Menu is shown
'   - A property page have been added for toolbar design
'
'=======================================================================================
'   I want specially thanks Jim Jose for his excellent McToolbar, I have used in my
'   jcToolbars some ideas from his usercontrol, such as chevrons and the way of loading
'   chevron picture as a separated window
'=======================================================================================

'=======================================================================================
'   There are still some unresolved problems (any help is wellcome):
'   - To modify used subclassing method to self subclassing in order to eliminate 2 class modules
'   - When button appears in picchevron (it is not visible) and you assign to this button
'     the function of unloading your program an error will occur.
'=======================================================================================


Option Explicit

'*************************************************************
'   Required Type Definitions
'*************************************************************
Private Type ToolbItem
    Caption As String
    Enabled As Boolean
    Key As String
    icon As StdPicture
    Iconsize As Integer
    BtnAlignment As AlignCont
    Tooltip As String
    BtnForeColor As OLE_COLOR
    font As StdFont
    Left As Long
    Top As Long
    Width As Long
    Height As Long
    R_Height As Long
    Type As jcBtnType
    State As jcBtnState
    Style As jcBtnStyle
    maskColor As OLE_COLOR
    UseMaskColor As Boolean
    Value As Boolean
End Type

Private Type TmpTBItem
    State As jcBtnState
    Value As Boolean
    icon As StdPicture
    Visible As Boolean
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type BITMAPINFOHEADER
    biSize          As Long
    biWidth         As Long
    biHeight        As Long
    biPlanes        As Integer
    biBitCount      As Integer
    biCompression   As Long
    biSizeImage     As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed       As Long
    biClrImportant  As Long
End Type

Private Type RGBTRIPLE
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBTRIPLE
End Type

'moving direction
'Public Enum MoveConst
'    ToLeft = 0
'    ToRight = 1
'End Enum

'Aligment icon and text
Public Enum AlignCont
    IconLeftTextRight = 0
    IconRightTextLeft = 1
    IconTopTextBottom = 2
    IconBottomTextTop = 3
End Enum

'type of toolbar item
Public Enum jcBtnType
    Button = 0
    Separator = 1
    EmptyButton = 2
End Enum

'button style
Public Enum jcBtnStyle
    [Normal button] = 0
    [Check button] = 1
    [Dropdown button] = 2
End Enum

'state constants
Public Enum jcBtnState
    STA_NORMAL = 0
    STA_OVER = 1
    STA_PRESSED = 2
    STA_OVERDOWN = 3
    STA_SELECTED = 4
    STA_DISABLED = 5
End Enum

'button property
Public Enum jcBtnChangeProp
    jcCaption = 1
    jcEnabled = 2
    jckey = 3
    jcIcon = 4
    jcIconSize = 5
    jcTooltip = 6
    jcBtnForeColor = 7
    jcStyle = 8
    jcState = 9
    jcValue = 10
    jcFont = 11
    jcAlignment = 12
    jcType = 13
    jcUseMaskColor = 14
    jcMaskColor = 15
End Enum

'gradient type
Public Enum jcGradConst
    VerticalGradient = 0
    HorizontalGradient = 1
    VCilinderGradient = 2
    HCilinderGradient = 3
End Enum

Private Type ICONINFO
    fIcon As Long
    xHotspot As Long
    yHotspot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Private Type BITMAP
   bmType       As Long
   bmWidth      As Long
   bmHeight     As Long
   bmWidthBytes As Long
   bmPlanes     As Integer
   bmBitsPixel  As Integer
   bmBits       As Long
End Type

'for bitmap conversion
Private Type Guid
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

'for bitmap conversion
Private Type pictDesc
   cbSizeofStruct As Long
   picType As Long
   hImage As Long
End Type

'*************************************************************
'   Constants
'*************************************************************
Private Const m_EmptyCaption As Integer = 16

'state constants
Private Const m_OffSet = 4

' Alignment constants
Private Const DT_CENTER = &H1
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_LEFT = &H0
Private Const TEXT_INACTIVE = &H808080

'*************************************************************
' Events
'*************************************************************
Public Event ButtonClick(btnIndex As Long, sKey As String, iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer, blnVisible As Boolean)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

'*************************************************************
' Members
'*************************************************************
Private m_ButtonCount As Long
Private ToolbarItem() As ToolbItem
Private TmpTBarItem() As TmpTBItem
Private m_BtnIndex As Long
Private m_PrevBtnIndex As Long
Private m_State As jcBtnState
Private R_Caption As RECT, R_Button As RECT
Private m_MinWidth As Integer
'Private m_MinHeight As Integer
Private ColorFrom As OLE_COLOR, ColorTo As OLE_COLOR
Private ColorFromOver As OLE_COLOR, ColorToOver As OLE_COLOR
Private ColorFromDown As OLE_COLOR, ColorToDown As OLE_COLOR
Private ColorToolbar As OLE_COLOR, ColorBorderPic As OLE_COLOR
Private ColorToRight As OLE_COLOR, ColorFromRight As OLE_COLOR
Private ColorToRightPress As OLE_COLOR, ColorFromRightPress As OLE_COLOR
Private ColorToRightOver As OLE_COLOR, ColorFromRightOver As OLE_COLOR
Private useMask As Boolean

Private m_SkipMouseMove As Boolean
Private m_TooltipShown As Boolean

'Public m_FontName As String
'Public m_Fontsize As Integer
'Public m_Bold As Boolean
'Public m_Italic As Boolean
'Public m_Underline As Boolean
'Public m_Strikethru As Boolean
'Public m_FontColor As Long

'*************************************************************
'   Required API Declarations
'*************************************************************
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpbi As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetNearestColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateIconIndirect Lib "user32" (piconinfo As ICONINFO) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lppictDesc As pictDesc, riid As Guid, ByVal fown As Long, ipic As IPicture) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long

Private Sub tmrTooltip_Timer()
    PicTB.ToolTipText = vbNullString
    m_TooltipShown = True
    tmrTooltip.Enabled = False
End Sub

'Private Sub UserControl_AmbientChanged(PropertyName As String)
'    If PropertyName = "BackColor" Then
'        UserControl.BackColor = Ambient.BackColor
'        DrawLeft
'        UserControl_Resize
'    End If
'End Sub

'==========================================================================
' Init, Read & Write UserControl
'==========================================================================
Private Sub UserControl_Initialize()
    SetDefaultThemeColor
    PicRight.BackColor = UserControl.BackColor
    PicLeft.BackColor = UserControl.BackColor
    PicTB.BackColor = UserControl.BackColor
    PicTB.Enabled = True
End Sub

Private Sub UserControl_InitProperties()
    UserControl.BackColor = Ambient.BackColor
    PicRight.BackColor = UserControl.BackColor
    PicLeft.BackColor = UserControl.BackColor
    PicTB.BackColor = UserControl.BackColor
    'Calculate_Size
    'Width = 400
End Sub

Private Sub UserControl_Resize()
Dim i As Integer
Dim MinHeight As Long
    
    MinHeight = MinimalHeight
    
    If UserControl.Height < MinHeight Then
        UserControl.Height = MinHeight
    End If
    
    If PicLeft.Height <> UserControl.ScaleHeight Then
        PicLeft.Height = UserControl.ScaleHeight
    End If
    
    PicTB.Move PicLeft.Left + PicLeft.Width, 0, Screen.Width \ Screen.TwipsPerPixelX, UserControl.ScaleHeight
    PicRight.Move UserControl.ScaleWidth - PicRight.Width, PicRight.Top, PicRight.Width, UserControl.ScaleHeight
    
    If m_ButtonCount > 0 Then
        
        For i = 1 To m_ButtonCount
            If ToolbarItem(i).Type = Button Then
                ToolbarItem(i).R_Height = PicTB.ScaleHeight - 2 * ToolbarItem(i).Top - 1
            Else
                ToolbarItem(i).R_Height = PicTB.ScaleHeight - 9
            End If
        Next i
        
    End If

    If Not Ambient.UserMode Then
        Calculate_Size
        InitialGradToolbar
        DrawTBtns
    End If
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Integer

    UserControl.BackColor = PropBag.ReadProperty("BackColor", Ambient.BackColor)
    m_ButtonCount = PropBag.ReadProperty("ButtonCount", 0)
    PicRight.BackColor = UserControl.Parent.BackColor 'UserControl.BackColor
    PicLeft.BackColor = UserControl.Parent.BackColor 'UserControl.BackColor
    PicTB.BackColor = UserControl.Parent.BackColor 'UserControl.BackColor
    PicTB.Enabled = PropBag.ReadProperty("Enabled", True)
    
    'Load toolButtons
    ReDim ToolbarItem(m_ButtonCount)
    ReDim TmpTBarItem(m_ButtonCount)
    
    For i = 1 To m_ButtonCount
        With ToolbarItem(i)
            .Caption = PropBag.ReadProperty("BtnCaption" & i, Empty)
            .Enabled = PropBag.ReadProperty("BtnEnabled" & i, True)
            Set .icon = PropBag.ReadProperty("BtnIcon" & i, Nothing)
            ConvertToIcon i
            .Iconsize = PropBag.ReadProperty("BtnIconSize" & i, 16)
            .Tooltip = PropBag.ReadProperty("BtnToolTipText" & i, vbNullString)
            .Key = PropBag.ReadProperty("BtnKey" & i, Empty)
            .BtnAlignment = PropBag.ReadProperty("BtnAlignment" & i, IconLeftTextRight)
            .Type = PropBag.ReadProperty("BtnType" & i, 0)
            .Style = PropBag.ReadProperty("BtnStyle" & i, 0)
            .BtnForeColor = PropBag.ReadProperty("BtnForeColor" & i, &H80000008)
            Set .font = PropBag.ReadProperty("BtnFont" & i, Ambient.font)
            .State = PropBag.ReadProperty("BtnState" & i, 0)
            .Value = PropBag.ReadProperty("BtnValue" & i, False)
            .UseMaskColor = PropBag.ReadProperty("BtnUseMaskColor" & i, True)
            .maskColor = PropBag.ReadProperty("BtnMaskColor" & i, QBColor(13))
            TmpTBarItem(i).State = .State
            TmpTBarItem(i).Value = .Value
            TmpTBarItem(i).Visible = True
        End With
    Next i
    
    SetDefaultThemeColor
    Calculate_Size
    DrawRight ColorFromRight, ColorToRight
    DrawLeft
    InitialGradToolbar
    DrawTBtns True
    m_BtnIndex = -1

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim i As Long

    PropBag.WriteProperty "BackColor", UserControl.BackColor, &HFFFFFF
    PropBag.WriteProperty "ButtonCount", m_ButtonCount, 0
    PropBag.WriteProperty "Enabled", PicTB.Enabled, True
    
    'Load toolButtons
    For i = 1 To m_ButtonCount
        With ToolbarItem(i)
            PropBag.WriteProperty "BtnCaption" & i, .Caption, Empty
            PropBag.WriteProperty "BtnEnabled" & i, .Enabled, True
            PropBag.WriteProperty "BtnIcon" & i, .icon, Nothing
            PropBag.WriteProperty "BtnIconSize" & i, .Iconsize, 16
            PropBag.WriteProperty "BtnToolTipText" & i, .Tooltip, vbNullString
            PropBag.WriteProperty "BtnKey" & i, .Key, Empty
            PropBag.WriteProperty "BtnAlignment" & i, .BtnAlignment, IconLeftTextRight
            PropBag.WriteProperty "BtnType" & i, .Type, 0
            PropBag.WriteProperty "BtnStyle" & i, .Style, 0
            PropBag.WriteProperty "BtnForeColor" & i, .BtnForeColor, &H80000008
            PropBag.WriteProperty "BtnFont" & i, .font, Ambient.font
            PropBag.WriteProperty "BtnState" & i, .State, 0
            PropBag.WriteProperty "BtnValue" & i, .Value, False
            PropBag.WriteProperty "BtnLeft" & i, .Left, 0
            PropBag.WriteProperty "BtnTop" & i, .Top, 0
            PropBag.WriteProperty "BtnWidth" & i, .Width, 0
            PropBag.WriteProperty "BtnHeight" & i, .Height, 0
            PropBag.WriteProperty "BtnUseMaskColor" & i, .UseMaskColor, True
            PropBag.WriteProperty "BtnMaskColor" & i, .maskColor, QBColor(13)
        End With
    Next i
    
End Sub

Private Sub PicTB_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If m_SkipMouseMove Then Exit Sub
    If m_State = STA_PRESSED Then Exit Sub
    
    m_BtnIndex = GetWhatButton(x, Y)
    tmrBtns.Enabled = CheckMouseOverPicTB
    
    If m_BtnIndex > -1 Then
        
        If TmpTBarItem(m_BtnIndex).Visible = False Then Exit Sub
        If ToolbarItem(m_BtnIndex).State = STA_DISABLED Then Exit Sub
        If ToolbarItem(m_BtnIndex).Type = Button Then
            
            If m_PrevBtnIndex <> m_BtnIndex Then
                m_PrevBtnIndex = m_BtnIndex
                PicTB.Cls
                If ToolbarItem(m_BtnIndex).Style = [Check button] Or ToolbarItem(m_BtnIndex).Style = [Dropdown button] Then
                    If TmpTBarItem(m_BtnIndex).Value Then
                        m_State = STA_OVERDOWN
                    Else
                        m_State = STA_OVER
                    End If
                Else
                    m_State = STA_OVER
                End If
                SetRect R_Button, ToolbarItem(m_BtnIndex).Left, ToolbarItem(m_BtnIndex).Top, ToolbarItem(m_BtnIndex).Width, ToolbarItem(m_BtnIndex).R_Height
            End If
            
            DrawBtn PicTB, m_State, R_Button, m_BtnIndex
            
            If m_TooltipShown = False Then
                PicTB.ToolTipText = ToolbarItem(m_BtnIndex).Tooltip
                tmrTooltip.Enabled = True
            Else
                m_TooltipShown = False
            End If
            
        End If
    Else
        tmrTooltip.Enabled = False
        PicTB.ToolTipText = vbNullString
        m_PrevBtnIndex = m_BtnIndex
        PicTB.Cls
        PicTB.Refresh
    End If
    
End Sub

Private Sub PicTB_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If Button = vbLeftButton Then
        
        If m_BtnIndex <> -1 Then
            If ToolbarItem(m_BtnIndex).Type = Separator Then Exit Sub
            If ToolbarItem(m_BtnIndex).State = STA_DISABLED Then Exit Sub
            m_State = STA_PRESSED
            DrawBtn PicTB, m_State, R_Button, m_BtnIndex
            m_State = STA_NORMAL
            m_SkipMouseMove = True
        Else
            DrawTBtns
        End If
        
    End If
    
End Sub

'==========================================================================
'  ButtonClick
'==========================================================================
Private Sub PicTB_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Button = vbLeftButton Then
        
        If m_BtnIndex <> -1 Then
            
            m_SkipMouseMove = False
            
            If ToolbarItem(m_BtnIndex).Type = Separator Then Exit Sub
            If ToolbarItem(m_BtnIndex).State = STA_DISABLED Then Exit Sub
            
            If ToolbarItem(m_BtnIndex).Style = [Check button] Or ToolbarItem(m_BtnIndex).Style = [Dropdown button] Then
                
                If TmpTBarItem(m_BtnIndex).Value Then
                    m_State = STA_NORMAL
                Else
                    m_State = STA_SELECTED
    '                blnDrop = True
                End If
                
                TmpTBarItem(m_BtnIndex).Value = Not TmpTBarItem(m_BtnIndex).Value
                TmpTBarItem(m_BtnIndex).State = m_State
                UpdateCheckValue PicTB, m_BtnIndex
                m_PrevBtnIndex = -1
                
            Else
                m_State = STA_OVER
                DrawBtn PicTB, m_State, R_Button, m_BtnIndex
                m_State = STA_NORMAL
            End If
            
    '        m_PrevBtn = Button
            RaiseEvent ButtonClick(m_BtnIndex, ToolbarItem(m_BtnIndex).Key, (ToolbarItem(m_BtnIndex).Left + PicTB.Left) * Screen.TwipsPerPixelX, (ToolbarItem(m_BtnIndex).Top) * Screen.TwipsPerPixelY, (ToolbarItem(m_BtnIndex).Width) * Screen.TwipsPerPixelX, (ToolbarItem(m_BtnIndex).R_Height) * Screen.TwipsPerPixelY, True)
            
        End If
    End If

End Sub

Private Sub tmrBtns_Timer()
    If CheckMouseOverPicTB = False Then
        PicTB.Cls
        PicTB.Refresh
    End If
End Sub

'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    UserControl.BackColor = New_BackColor
'    PicRight.BackColor = New_BackColor
'    PicLeft.BackColor = New_BackColor
'    PropertyChanged "BackColor"
'    DrawTBtns
'End Property

Public Property Get ButtonCount() As Integer
    ButtonCount = m_ButtonCount
End Property

Public Property Get BtnCaption(ByVal Index As Integer) As String
    BtnCaption = ToolbarItem(Index).Caption
End Property

Public Property Let BtnCaption(ByVal Index As Integer, ByVal New_Value As String)
    
    If ToolbarItem(Index).Caption <> New_Value Then
        ToolbarItem(Index).Caption = New_Value
        ChangeBtnProperty jcCaption, Index, New_Value
        PropertyChanged "BtnCaption"
    End If
    
End Property

Public Property Get BtnEnabled(ByVal Index As Integer) As Boolean
    BtnEnabled = ToolbarItem(Index).Enabled
End Property

Public Property Let BtnEnabled(ByVal Index As Integer, ByVal New_Value As Boolean)
    
    If ToolbarItem(Index).Enabled <> New_Value Then
        ToolbarItem(Index).Enabled = New_Value
        ChangeBtnProperty jcEnabled, Index, New_Value
        PropertyChanged "BtnEnabled"
    End If
    
End Property

'Public Property Get BtnIcon(ByVal Index As Integer) As StdPicture
'    Set BtnIcon = ToolbarItem(Index).icon
'End Property

'Public Property Set BtnIcon(ByVal Index As Integer, ByVal New_Picture As StdPicture)
'    Set ToolbarItem(Index).icon = New_Picture
'    ChangeBtnProperty jcIcon, Index, New_Picture
'    PropertyChanged "BtnIcon"
'End Property
'
'Public Property Get BtnIconSize(ByVal Index As Integer) As Integer
'    BtnIconSize = ToolbarItem(Index).Iconsize
'End Property

'Public Property Let BtnIconSize(ByVal Index As Integer, ByVal New_Value As Integer)
'    If ToolbarItem(Index).Iconsize = New_Value Then Exit Property
'    ToolbarItem(Index).Iconsize = New_Value
'    ChangeBtnProperty jcIconSize, Index, New_Value
'    PropertyChanged "BtnIconSize"
'End Property

'Public Property Get BtnToolTipText(ByVal Index As Integer) As String
'    BtnToolTipText = ToolbarItem(Index).Tooltip
'End Property

Public Property Let BtnToolTipText(ByVal Index As Integer, ByVal New_Value As String)
    
    If ToolbarItem(Index).Tooltip <> New_Value Then
        ToolbarItem(Index).Tooltip = New_Value
        ChangeBtnProperty jcTooltip, Index, New_Value
        PropertyChanged "BtnToolTipText"
    End If
    
End Property

'Public Property Get BtnKey(ByVal Index As Integer) As String
'    BtnKey = ToolbarItem(Index).Key
'End Property
'
'Public Property Let BtnKey(ByVal Index As Integer, ByVal New_Value As String)
'    If ToolbarItem(Index).Key = New_Value Then Exit Property
'    ToolbarItem(Index).Key = New_Value
'    ChangeBtnProperty jckey, Index, New_Value
'    PropertyChanged "BtnKey"
'End Property

'Public Property Get BtnAlignment(ByVal Index As Integer) As AlignCont
'    BtnAlignment = ToolbarItem(Index).BtnAlignment
'End Property
'
'Public Property Let BtnAlignment(ByVal Index As Integer, ByVal New_Value As AlignCont)
'    If ToolbarItem(Index).BtnAlignment = New_Value Then Exit Property
'    ToolbarItem(Index).BtnAlignment = New_Value
'    ChangeBtnProperty jcAlignment, Index, New_Value
'    PropertyChanged "BtnAlignment"
'End Property

'Public Property Get BtnType(ByVal Index As Integer) As jcBtnType
'    BtnType = ToolbarItem(Index).Type
'End Property
'
'Public Property Let BtnType(ByVal Index As Integer, ByVal New_Value As jcBtnType)
'    If ToolbarItem(Index).Type = New_Value Then Exit Property
'    ToolbarItem(Index).Type = New_Value
'    ChangeBtnProperty jcType, Index, New_Value
'    PropertyChanged "BtnType"
'End Property

'Public Property Get BtnStyle(ByVal Index As Integer) As jcBtnStyle
'    BtnStyle = ToolbarItem(Index).Style
'End Property
'
'Public Property Let BtnStyle(ByVal Index As Integer, ByVal New_Value As jcBtnStyle)
'    If ToolbarItem(Index).Style = New_Value Then Exit Property
'    ToolbarItem(Index).Style = New_Value
'    ChangeBtnProperty jcStyle, Index, New_Value
'    PropertyChanged "BtnStyle"
'End Property

'Public Property Get BtnForeColor(ByVal Index As Integer) As OLE_COLOR
'    BtnForeColor = ToolbarItem(Index).BtnForeColor
'End Property
'
'Public Property Let BtnForeColor(ByVal Index As Integer, ByVal New_Value As OLE_COLOR)
'    If ToolbarItem(Index).BtnForeColor = New_Value Then Exit Property
'    ToolbarItem(Index).BtnForeColor = New_Value
'    ChangeBtnProperty jcBtnForeColor, Index, New_Value
'    PropertyChanged "BtnForeColor"
'End Property
'
'Public Property Get BtnFont(ByVal Index As Integer) As font
'    Set BtnFont = ToolbarItem(Index).font
'End Property

'Public Property Let BtnFont(ByVal Index As Integer, ByVal New_Value As font)
'    If ToolbarItem(Index).font.Name = New_Value.Name And _
'    ToolbarItem(Index).font.Bold = New_Value.Bold And _
'    ToolbarItem(Index).font.Italic = New_Value.Italic And _
'    ToolbarItem(Index).font.SIZE = New_Value.SIZE And _
'    ToolbarItem(Index).font.Strikethrough = New_Value.Strikethrough And _
'    ToolbarItem(Index).font.Underline = New_Value.Underline And _
'    ToolbarItem(Index).font.Weight = New_Value.Weight Then Exit Property
'    Set ToolbarItem(Index).font = New_Value
'    ChangeBtnProperty jcFont, Index, New_Value
'    PropertyChanged "BtnFont"
'End Property

Public Property Get BtnState(ByVal Index As Integer) As Integer
    BtnState = ToolbarItem(Index).State
End Property

Public Property Let BtnState(ByVal Index As Integer, ByVal New_Value As Integer)
    
    If ToolbarItem(Index).State <> New_Value Then
        ToolbarItem(Index).State = New_Value
        TmpTBarItem(Index).State = New_Value
        ChangeBtnProperty jcState, Index, New_Value
        PropertyChanged "BtnState"
    End If
    
End Property

Public Property Get BtnValue(ByVal Index As Integer) As Boolean
    BtnValue = ToolbarItem(Index).Value
End Property

Public Property Let BtnValue(ByVal Index As Integer, ByVal New_Value As Boolean)
    
    If ToolbarItem(Index).Value <> New_Value Then
        ToolbarItem(Index).Value = New_Value
        TmpTBarItem(Index).Value = New_Value
        ChangeBtnProperty jcValue, Index, New_Value
        PropertyChanged "BtnValue"
    End If
    
End Property

'Public Property Get BtnLeft(ByVal Index As Integer) As Integer
'    BtnLeft = ToolbarItem(Index).Left
'End Property
'
'Public Property Let BtnLeft(ByVal Index As Integer, ByVal New_Value As Integer)
'    ToolbarItem(Index).Left = New_Value
'    PropertyChanged "BtnLeft"
'End Property
'
'Public Property Get BtnTop(ByVal Index As Integer) As Integer
'    BtnTop = ToolbarItem(Index).Top
'End Property
'
'Public Property Let BtnTop(ByVal Index As Integer, ByVal New_Value As Integer)
'    ToolbarItem(Index).Top = New_Value
'    PropertyChanged "BtnTop"
'End Property

Public Property Get BtnWidth(ByVal Index As Integer) As Integer
    BtnWidth = ToolbarItem(Index).Width
End Property

Public Property Let BtnWidth(ByVal Index As Integer, ByVal New_Value As Integer)
    If ToolbarItem(Index).Width <> New_Value Then
        ToolbarItem(Index).Width = New_Value
        PropertyChanged "BtnWidth"
    End If
End Property

'Public Property Get btnHeight(ByVal Index As Integer) As Integer
'    btnHeight = ToolbarItem(Index).Height
'End Property
'
'Public Property Let btnHeight(ByVal Index As Integer, ByVal New_Value As Integer)
'    ToolbarItem(Index).Height = New_Value
'    PropertyChanged "BtnHeight"
'End Property

'Public Property Get BtnRHeight(ByVal Index As Integer) As Integer
'    BtnRHeight = ToolbarItem(Index).R_Height
'End Property
'
'Public Property Let BtnRHeight(ByVal Index As Integer, ByVal New_Value As Integer)
'    ToolbarItem(Index).R_Height = New_Value
'    PropertyChanged "BtnRHeight"
'End Property
'
'Public Property Get BtnUseMaskColor(ByVal Index As Integer) As Boolean
'    BtnUseMaskColor = ToolbarItem(Index).UseMaskColor
'End Property
'
'Public Property Let BtnUseMaskColor(ByVal Index As Integer, ByVal New_Value As Boolean)
'    If ToolbarItem(Index).UseMaskColor = New_Value Then Exit Property
'    ToolbarItem(Index).UseMaskColor = New_Value
'    ChangeBtnProperty jcUseMaskColor, Index, New_Value
'    PropertyChanged "BtnUseMaskColor"
'End Property

'Public Property Get BtnMaskColor(ByVal Index As Integer) As OLE_COLOR
'    BtnMaskColor = ToolbarItem(Index).maskColor
'End Property
'
'Public Property Let BtnMaskColor(ByVal Index As Integer, ByVal New_Value As OLE_COLOR)
'    If ToolbarItem(Index).maskColor = New_Value Then Exit Property
'    ToolbarItem(Index).maskColor = New_Value
'    ChangeBtnProperty jcMaskColor, Index, New_Value
'    PropertyChanged "BtnMaskColor"
'End Property

Public Property Get Enabled() As Boolean
    Enabled = PicTB.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    
    If PicTB.Enabled <> New_Enabled Then
        PicTB.Enabled = New_Enabled
        PropertyChanged "Enabled"
    End If
    
End Property

Private Function CheckMouseOverPicTB() As Boolean
Dim pt As POINTAPI

    GetCursorPos pt
    CheckMouseOverPicTB = (WindowFromPoint(pt.x, pt.Y) = PicTB.hWnd)
    
End Function

Private Sub DrawRight(FromColor As Long, ToColor As Long)
Dim r As RECT, lColor As Long

    SetRect r, 0, 0, 2, PicRight.Height
    DrawVGradientEx PicRight.hDC, ColorTo, ColorFrom, r.Left, r.Top, r.Right, r.Bottom
    
    SetRect r, 2, 0, PicRight.Width, PicRight.Height
    DrawVGradientEx PicRight.hDC, FromColor, ToColor, r.Left, r.Top, r.Right, r.Bottom
    
    lColor = TranslateColor(Ambient.BackColor)
    SetPixel PicRight.hDC, PicRight.Width - 1, 0, lColor
    SetPixel PicRight.hDC, PicRight.Width - 1, PicRight.Height - 1, lColor
    
    lColor = BlendColors(TranslateColor(Ambient.BackColor), FromColor)
    SetPixel PicRight.hDC, PicRight.Width - 2, 0, lColor
    SetPixel PicRight.hDC, PicRight.Width - 1, 1, lColor
    
    lColor = BlendColors(TranslateColor(ColorTo), FromColor)
    SetPixel PicRight.hDC, 0, 0, lColor
    SetPixel PicRight.hDC, 1, 1, lColor
    SetPixel PicRight.hDC, 1, 0, FromColor
    
    lColor = BlendColors(TranslateColor(Ambient.BackColor), ToColor)
    SetPixel PicRight.hDC, PicRight.Width - 2, PicRight.Height - 1, lColor
    SetPixel PicRight.hDC, PicRight.Width - 1, PicRight.Height - 2, lColor
    
    lColor = BlendColors(TranslateColor(ColorFrom), ToColor)
    SetPixel PicRight.hDC, 0, PicRight.Height - 1, lColor
    SetPixel PicRight.hDC, 1, PicRight.Height - 2, lColor
    SetPixel PicRight.hDC, 1, PicRight.Height - 1, ToColor
            
    PicRight.Refresh
    
End Sub

Private Sub DrawLeft()
Dim r As RECT
Dim lColor As Long
Dim i As Long
Dim yTop As Long
Dim NumRect As Long
    
    SetRect r, 0, 0, PicLeft.Width, PicLeft.Height
    DrawVGradientEx PicLeft.hDC, ColorTo, ColorFrom, r.Left, r.Top, r.Right, r.Bottom

    SetRect r, 2, PicLeft.Height - 1, PicLeft.Width, PicLeft.Height - 1
    APILineEx PicLeft.hDC, r.Left, r.Top, r.Right, r.Bottom, ColorToolbar

    lColor = TranslateColor(Ambient.BackColor)
    SetPixel PicLeft.hDC, 0, 0, lColor
    SetPixel PicLeft.hDC, 0, PicRight.Height - 1, lColor
    SetPixel PicLeft.hDC, 0, PicRight.Height - 2, lColor
    SetPixel PicLeft.hDC, 1, PicRight.Height - 1, lColor

    lColor = BlendColors(vbWhite, ColorTo)
    SetPixel PicLeft.hDC, 1, 0, lColor
    SetPixel PicLeft.hDC, 0, 1, lColor

    lColor = BlendColors(ColorBorderPic, ColorFrom)
    SetPixel PicLeft.hDC, 1, PicRight.Height - 3, lColor
    
    lColor = BlendColors(TranslateColor(Ambient.BackColor), ColorFrom)
    SetPixel PicLeft.hDC, 0, PicRight.Height - 3, lColor
    SetPixel PicLeft.hDC, 1, PicRight.Height - 2, lColor
    
    NumRect = (PicRight.ScaleHeight - PicRight.ScaleHeight * 0.4) \ 4
    yTop = (PicRight.ScaleHeight - 4 * (NumRect - 1) - 1) \ 2
    
    For i = 0 To NumRect - 1
        SetRect r, 5, yTop + 4 * i, 1, 1
        ApiRectangle PicLeft.hDC, r.Left, r.Top, r.Right, r.Bottom, vbWhite
        SetRect r, 4, (yTop - 1) + 4 * i, 1, 1
        ApiRectangle PicLeft.hDC, r.Left, r.Top, r.Right, r.Bottom, ColorToolbar
    Next i
    
    PicLeft.Refresh
    
End Sub

Private Sub SetDefaultThemeColor()

    ColorFromOver = RGB(216, 216, 216)
    ColorToOver = RGB(242, 242, 242)
    ColorFromDown = RGB(184, 184, 184)
    ColorToDown = ColorFromOver
    ColorFromRightOver = RGB(246, 246, 246)
    ColorToRightOver = RGB(208, 208, 208)
    ColorFromRightPress = RGB(188, 188, 188)
    ColorToRightPress = ColorFromOver
    
    ColorFrom = ColorFromDown
    ColorTo = ColorToOver
    ColorToolbar = RGB(154, 154, 154)
    ColorBorderPic = RGB(110, 110, 110)
    ColorFromRight = ColorToRightOver
    ColorToRight = RGB(146, 146, 146)
    
End Sub

'Public Function AddButton(Optional ByVal m_Type As jcBtnType = Button, Optional ByVal m_Caption As String = Empty, Optional ByVal m_Key As String = Empty, Optional ByVal m_Icon As StdPicture = Nothing, Optional ByVal m_IconSize As Integer = 16, Optional ByVal m_BtnAlignment As AlignCont = IconLeftTextRight, Optional ByVal m_Tooltip As String = Empty, Optional ByVal m_ForeColor As OLE_COLOR = &H80000008, Optional ByVal m_Font As StdFont, Optional ByVal m_Style As jcBtnStyle = [Normal button], Optional ByVal m_State As jcBtnState = STA_NORMAL, Optional ByVal m_UseMaskColor As Boolean = True, Optional ByVal m_MaskColor As OLE_COLOR = vbMagenta) As Integer
'
''    Dim i As Integer, ix As Integer, iy As Integer, iw As Integer, ih As Integer
'
'    Dim X As New StdFont
'    X.SIZE = 8
'    X.Name = "MS Sans Serif"
'
'    m_ButtonCount = m_ButtonCount + 1
'
'    ReDim Preserve ToolbarItem(m_ButtonCount)
'    ReDim Preserve TmpTBarItem(m_ButtonCount)
'
'    ToolbarItem(m_ButtonCount).Type = m_Type
'
'    If Not (m_Font Is Nothing) Then
'        Set ToolbarItem(m_ButtonCount).font = m_Font
'    Else
'        Set ToolbarItem(m_ButtonCount).font = X
'    End If
'
'    ToolbarItem(m_ButtonCount).Caption = m_Caption
'    ToolbarItem(m_ButtonCount).Key = m_Key
'
'    If Not (m_Icon Is Nothing) Then
'        Set ToolbarItem(m_ButtonCount).icon = m_Icon
'        ConvertToIcon CInt(m_ButtonCount)
'    Else
'        Set ToolbarItem(m_ButtonCount).icon = Nothing
'    End If
'
'    ToolbarItem(m_ButtonCount).Iconsize = m_IconSize
'    ToolbarItem(m_ButtonCount).BtnAlignment = m_BtnAlignment
'    ToolbarItem(m_ButtonCount).Tooltip = m_Tooltip
'    ToolbarItem(m_ButtonCount).BtnForeColor = m_ForeColor
'    ToolbarItem(m_ButtonCount).Style = m_Style
'    ToolbarItem(m_ButtonCount).State = m_State
'    ToolbarItem(m_ButtonCount).UseMaskColor = m_UseMaskColor
'    ToolbarItem(m_ButtonCount).maskColor = m_MaskColor
'
'    AddButton = m_ButtonCount
'    Calculate_Size
'
'    If Ambient.UserMode Then
'        'Height = MinimalHeight
'        InitialGradToolbar
'    End If
'
'    Width = m_MinWidth
'
'End Function

Private Sub DrawSeparator(ColorLine As OLE_COLOR, lLeft As Long, lTop As Long, lHeight As Long)
    APILineEx PicTB.hDC, lLeft + 1, lTop, lLeft + 1, lTop + lHeight, ColorLine
    APILineEx PicTB.hDC, lLeft + 2, lTop + 1, lLeft + 2, lTop + lHeight + 1, vbWhite
End Sub

Public Sub ChangeBtnProperty(intOption As jcBtnChangeProp, intI As Integer, NewValue As Variant)
    Select Case intOption
        Case jcAlignment
            ToolbarItem(intI).BtnAlignment = NewValue
            Calculate_Size 'intI
            Width = m_MinWidth
            InitialGradToolbar
            DrawTBtns False
        Case jcCaption
            ToolbarItem(intI).Caption = NewValue
            Calculate_Size intI
            Width = m_MinWidth
            DrawTBtns True, CLng(intI)
        Case jcEnabled
            ToolbarItem(intI).Enabled = NewValue
            ToolbarItem(intI).State = STA_DISABLED
            TmpTBarItem(intI).State = STA_DISABLED
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jckey
            ToolbarItem(intI).Key = NewValue
        Case jcIcon
            Set ToolbarItem(intI).icon = NewValue
            ConvertToIcon intI
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jcIconSize
            ToolbarItem(intI).Iconsize = NewValue
            Calculate_Size intI, True
            Width = m_MinWidth
            InitialGradToolbar
            DrawTBtns False
        Case jcTooltip
            ToolbarItem(intI).Tooltip = NewValue
        Case jcBtnForeColor
            ToolbarItem(intI).BtnForeColor = NewValue
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jcStyle
            ToolbarItem(intI).Style = NewValue
            Calculate_Size intI
            Width = m_MinWidth
            DrawTBtns True, CLng(intI), m_ButtonCount
        Case jcState
            ToolbarItem(intI).State = NewValue
            TmpTBarItem(intI).State = NewValue
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jcValue
            If NewValue = True Then
                ToolbarItem(intI).State = STA_PRESSED
                TmpTBarItem(intI).State = STA_PRESSED
            Else
                ToolbarItem(intI).State = STA_NORMAL
                TmpTBarItem(intI).State = STA_NORMAL
            End If
            ToolbarItem(intI).Value = NewValue
            TmpTBarItem(intI).Value = NewValue
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jcFont
            Set ToolbarItem(intI).font = NewValue
            Calculate_Size intI
            Width = m_MinWidth
            InitialGradToolbar
            DrawTBtns False
        Case jcType
            ToolbarItem(intI).Type = NewValue
            Calculate_Size intI
            Width = m_MinWidth
            InitialGradToolbar
            DrawTBtns False
        Case jcUseMaskColor
            ToolbarItem(intI).UseMaskColor = NewValue
            DrawTBtns True, CLng(intI), CLng(intI)
        Case jcMaskColor
            ToolbarItem(intI).maskColor = NewValue
            DrawTBtns True, CLng(intI), CLng(intI)
    End Select
End Sub

'Public Function DeleteButton(ByVal Index As Integer) As Integer
'    'Deletes an existing button from the control.
'    Dim i As Long, BtnWidth As Long
'
'    If Index < 0 Or Index > m_ButtonCount Then
'        DeleteButton = -1
'        Exit Function
'    End If
'
'    If m_ButtonCount = 0 Then Exit Function
'    BtnWidth = ToolbarItem(Index).Width + m_OffSet / 2
'
'    For i = Index To m_ButtonCount - 1
'        ToolbarItem(i) = ToolbarItem(i + 1)
'        TmpTBarItem(i) = TmpTBarItem(i + 1)
'        ToolbarItem(i).Left = ToolbarItem(i).Left - BtnWidth
'    Next i
'
'    m_ButtonCount = m_ButtonCount - 1
'    ReDim Preserve ToolbarItem(m_ButtonCount)
'    ReDim Preserve TmpTBarItem(m_ButtonCount)
'
'    'initial toolbar width
'    UserControl.Width = (ToolbarItem(m_ButtonCount).Left + ToolbarItem(m_ButtonCount).Width + m_OffSet / 2 + PicLeft.Width + PicRight.Width) * 15
'    Calculate_Size
'    Width = m_MinWidth
'    Height = m_MinHeight
'    InitialGradToolbar
'    DrawTBtns True
'    DeleteButton = m_ButtonCount
'End Function
'
'Public Function MoveButton(ByVal Index As Integer, lDirection As MoveConst) As Integer
'    Dim ToolbarItem_Aux As ToolbItem, NewIndex As Integer
'    Dim TmpTBarItem_Aux As TmpTBItem
''    Dim i As Long
'    Dim BtnWidth As Long, iFrom As Long, iTo As Long
'
'    Select Case lDirection
'        Case ToLeft
'            If Index > 1 Then NewIndex = Index - 1 Else Exit Function
'            BtnWidth = ToolbarItem(Index).Width + m_OffSet / 2
'            ToolbarItem(Index).Left = ToolbarItem(NewIndex).Left
'            ToolbarItem(NewIndex).Left = ToolbarItem(Index).Left + BtnWidth
'            iFrom = NewIndex
'            iTo = Index
'        Case ToRight
'            If Index < m_ButtonCount Then NewIndex = Index + 1 Else Exit Function
'            BtnWidth = ToolbarItem(NewIndex).Width + m_OffSet / 2
'            ToolbarItem(NewIndex).Left = ToolbarItem(Index).Left
'            ToolbarItem(Index).Left = ToolbarItem(NewIndex).Left + BtnWidth
'            iFrom = Index
'            iTo = NewIndex
'    End Select
'
'    ToolbarItem_Aux = ToolbarItem(NewIndex)
'    ToolbarItem(NewIndex) = ToolbarItem(Index)
'    ToolbarItem(Index) = ToolbarItem_Aux
'
'    TmpTBarItem_Aux = TmpTBarItem(NewIndex)
'    TmpTBarItem(NewIndex) = TmpTBarItem(Index)
'    TmpTBarItem(Index) = TmpTBarItem_Aux
'
'    DrawTBtns True, iFrom, iTo
'    MoveButton = NewIndex
'
'End Function

Private Sub Calculate_Size(Optional j As Integer = 1, Optional blnSize As Boolean = False)
Dim i As Integer ', k As Integer, MaxHeight As Integer
Dim ix As Long, iy As Long, iw As Long, ih As Long
Dim x As New StdFont

    x.SIZE = 8
    x.name = "Tahoma"
    
    m_MinWidth = 400

    If m_ButtonCount = 0 Then Exit Sub
    
    For i = j To m_ButtonCount
        If i > 0 Then
            ix = ToolbarItem(i - 1).Left + ToolbarItem(i - 1).Width + m_OffSet \ 2
        Else
            ix = ToolbarItem(i + 1).Left
        End If
        
        If ToolbarItem(i).Type = Button Then
            iy = 2
            
            If Not (ToolbarItem(i).font Is Nothing) Then
                Set UserControl.font = ToolbarItem(i).font
            Else
                Set UserControl.font = x
            End If
            
            If ToolbarItem(i).icon Is Nothing Then  'there is no icon
                If LenB(ToolbarItem(i).Caption) <> 0 Then   'there is no caption
                    iw = 2 * m_OffSet + TextWidth(ToolbarItem(i).Caption)
                    ih = TextHeight(ToolbarItem(i).Caption) + m_OffSet * 2
                Else
                    iw = 2 * m_OffSet + m_EmptyCaption
                    ih = m_EmptyCaption + m_OffSet * 2
                End If
            Else
                If LenB(ToolbarItem(i).Caption) <> 0 Then   'there is no caption
                    Select Case ToolbarItem(i).BtnAlignment
                        Case IconLeftTextRight, IconRightTextLeft
                            iw = 3 * m_OffSet + TextWidth(ToolbarItem(i).Caption) + ToolbarItem(i).Iconsize
                            If TextHeight(ToolbarItem(i).Caption) > ToolbarItem(i).Iconsize Then
                                ih = TextHeight(ToolbarItem(i).Caption) + m_OffSet * 2
                            Else
                                ih = ToolbarItem(i).Iconsize + m_OffSet * 2
                            End If
                        Case IconTopTextBottom, IconBottomTextTop
                            If TextWidth(ToolbarItem(i).Caption) > ToolbarItem(i).Iconsize Then
                                iw = TextWidth(ToolbarItem(i).Caption) + m_OffSet * 2
                            Else
                                iw = ToolbarItem(i).Iconsize + m_OffSet * 2
                            End If
                            ih = m_OffSet * 3 + TextHeight(ToolbarItem(i).Caption) + ToolbarItem(i).Iconsize + 1
                    End Select
                Else
                    iw = m_OffSet * 2 + ToolbarItem(i).Iconsize
                    ih = ToolbarItem(i).Iconsize + m_OffSet * 2
                End If
            End If
            If ToolbarItem(i).Style = [Dropdown button] Then iw = iw + 13
        Else    'it is separator
            iy = 4
            iw = 2
            ih = PicTB.ScaleHeight - 9
            ToolbarItem(i).R_Height = ih
        End If
        
        ToolbarItem(i).Left = ix
        ToolbarItem(i).Top = iy
        ToolbarItem(i).Width = iw
        ToolbarItem(i).Height = ih
        
        If ToolbarItem(i).Type = Button Then
            ToolbarItem(i).R_Height = PicTB.ScaleHeight - ToolbarItem(i).Top * 2 - 1
        Else
            ToolbarItem(i).R_Height = ih
        End If
    Next i
    
    'initial toolbar width
    m_MinWidth = (ix + iw + m_OffSet / 2 + PicLeft.Width + PicRight.Width) * Screen.TwipsPerPixelX
    
    If blnSize Then
        Width = m_MinWidth
        'UserControl_Resize
    End If

End Sub

Private Function GetWhatButton(ByVal x As Integer, ByVal Y As Integer) As Integer
Dim i As Integer

    For i = 0 To m_ButtonCount
        If TmpTBarItem(i).Visible = True And ToolbarItem(i).Type = Button Then
            If x > ToolbarItem(i).Left And x < ToolbarItem(i).Left + ToolbarItem(i).Width Then
                If Y > ToolbarItem(i).Top And Y < ToolbarItem(i).Top + ToolbarItem(i).R_Height Then
                    GetWhatButton = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    GetWhatButton = -1
    
End Function

' full version of APILine
Private Sub APILineEx(lhdcEx As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lColor As Long)
Dim hPen As Long, hPenOld As Long

    'Use the API LineTo for Fast Drawing
    
    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(lhdcEx, hPen)
    MoveToEx lhdcEx, X1, Y1, 0
    LineTo lhdcEx, X2, Y2
    SelectObject lhdcEx, hPenOld
    DeleteObject hPen
    DeleteObject hPenOld
    
End Sub

Private Function ApiRectangle(ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal w As Long, ByVal H As Long, Optional lColor As OLE_COLOR = -1) As Long
Dim hPen As Long, hPenOld As Long

    hPen = CreatePen(0, 1, lColor)
    hPenOld = SelectObject(hDC, hPen)
    MoveToEx hDC, x, Y, 0
    LineTo hDC, x + w, Y
    LineTo hDC, x + w, Y + H
    LineTo hDC, x, Y + H
    LineTo hDC, x, Y
    SelectObject hDC, hPenOld
    DeleteObject hPen
    DeleteObject hPenOld
    
End Function

Private Sub DrawVGradientEx(lhdcEx As Long, lEndColor As Long, lStartcolor As Long, ByVal x As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
Dim dR As Single, dG As Single, dB As Single
Dim sR As Single, sG As Single, sB As Single
Dim eR As Single, eG As Single, eB As Single
Dim ni As Long

    'Draw a Vertical Gradient in the current HDC
    
    If Y2 = 0 Then Exit Sub
    
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / Y2
    dG = (sG - eG) / Y2
    dB = (sB - eB) / Y2
    
    For ni = 0 To Y2
        APILineEx lhdcEx, x, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next ni
    
End Sub

Private Sub DrawGradientEx(lhdcEx As Long, lEndColor As Long, lStartcolor As Long, ByVal x As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional blnVertical As Boolean = True)
Dim dR As Single, dG As Single, dB As Single
Dim sR As Single, sG As Single, sB As Single
Dim eR As Single, eG As Single, eB As Single
Dim ni As Long
    
    If Y2 = 0 Then Exit Sub
    If X2 = 0 Then Exit Sub
    
    'Draw a Vertical or horizontal Gradient in the current HDC
    
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    
    If blnVertical Then
    
        dR = (sR - eR) / Y2
        dG = (sG - eG) / Y2
        dB = (sB - eB) / Y2
        
        For ni = 1 To Y2 - 1
            APILineEx lhdcEx, x, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next ni
        
    Else
        
        dR = (sR - eR) / X2
        dG = (sG - eG) / X2
        dB = (sB - eB) / X2
        
        For ni = 1 To X2 - 1
            APILineEx lhdcEx, x + ni, Y, x + ni, Y2, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
        Next ni
        
    End If
    
End Sub

'Blend two colors
Private Function BlendColors(ByVal lcolor1 As Long, ByVal lcolor2 As Long) As Long
    
    On Error GoTo Hell
    BlendColors = RGB(((lcolor1 And &HFF) + (lcolor2 And &HFF)) / 2, (((lcolor1 \ &H100) And &HFF) + ((lcolor2 \ &H100) And &HFF)) / 2, (((lcolor1 \ &H10000) And &HFF) + ((lcolor2 \ &H10000) And &HFF)) / 2)

Hell:
End Function

'System color code to long rgb
Private Function TranslateColor(ByVal lColor As Long) As Long
    If OleTranslateColor(lColor, 0, TranslateColor) Then
          TranslateColor = -1
    End If
End Function

Private Sub DrawGradientInRectangle(lhdcEx As Long, lStartcolor As Long, lEndColor As Long, r As RECT, GradientType As jcGradConst, Optional blnDrawBorder As Boolean = False, Optional lBorderColor As Long = vbBlack, Optional LightCenter As Double = 2.01)
    
    Select Case GradientType
        Case VerticalGradient
            DrawGradientEx lhdcEx, lEndColor, lStartcolor, r.Left, r.Top, r.Right + r.Left, r.Bottom, True
        Case HorizontalGradient
            DrawGradientEx lhdcEx, lEndColor, lStartcolor, r.Left, r.Top, r.Right, r.Bottom + r.Top, False
        Case VCilinderGradient
            DrawGradCilinder lhdcEx, lStartcolor, lEndColor, r, True, LightCenter
        Case HCilinderGradient
            DrawGradCilinder lhdcEx, lStartcolor, lEndColor, r, False, LightCenter
    End Select
    
    If blnDrawBorder Then ApiRectangle lhdcEx, r.Left, r.Top, r.Right, r.Bottom, lBorderColor
    
End Sub

Private Sub DrawGradCilinder(lhdcEx As Long, lStartcolor As Long, lEndColor As Long, r As RECT, Optional ByVal blnVertical As Boolean = True, Optional ByVal LightCenter As Double = 2.01)
    
    If LightCenter <= 1# Then LightCenter = 1.01
    
    If blnVertical Then
        DrawGradientEx lhdcEx, lStartcolor, lEndColor, r.Left, r.Top, r.Right + r.Left, r.Bottom / LightCenter, True
        DrawGradientEx lhdcEx, lEndColor, lStartcolor, r.Left, r.Top + r.Bottom / LightCenter - 1, r.Right + r.Left, (LightCenter - 1) * r.Bottom / LightCenter + 1, True
    Else
        DrawGradientEx lhdcEx, lStartcolor, lEndColor, r.Left, r.Top, r.Right / LightCenter, r.Bottom + r.Top, False
        DrawGradientEx lhdcEx, lEndColor, lStartcolor, r.Left + r.Right / LightCenter - 1, r.Top, (LightCenter - 1) * r.Right / LightCenter + 1, r.Bottom + r.Top, False
    End If
    
End Sub

Private Sub DrawCaption(Pic As PictureBox, sText As String, FntColor As Long, RCaption As RECT, MyFont As StdFont, Optional BlnCenter As Boolean = False)
Dim textAligment As Long
    
    Pic.ForeColor = FntColor
    Set Pic.font = MyFont
    
    textAligment = DT_VCENTER Or DT_SINGLELINE
    
    'Set the rectangle's values
    If BlnCenter = False Then
        textAligment = textAligment Or DT_LEFT
    Else
        textAligment = textAligment Or DT_CENTER
    End If
    
    'Draw text in PicTB or picChevron
    DrawTextEx Pic.hDC, sText, LenB(sText), RCaption, textAligment, ByVal 0&
    
End Sub

'Drawing buttons in picTB
Private Sub DrawTBtns(Optional blnClear As Boolean = False, Optional iFrom As Long = 1, Optional iTo As Long = -1)
Dim i As Long, R1 As RECT
Dim r As RECT
    
    If m_ButtonCount = 0 Then Exit Sub
    
    If iTo = -1 Then iTo = m_ButtonCount
    If iFrom < 1 Then iFrom = 1
    If iTo > m_ButtonCount Then iTo = m_ButtonCount
    
    'clearing picTB background
    If blnClear Then
        
        If iFrom = m_ButtonCount Then
            SetRect r, ToolbarItem(iFrom).Left, 0, PicTB.ScaleWidth - ToolbarItem(iFrom).Left - 2, PicTB.ScaleHeight - 2
        Else
            SetRect r, ToolbarItem(iFrom).Left - 1, 0, ToolbarItem(iTo).Left - ToolbarItem(iFrom).Left + ToolbarItem(iTo).Width + m_OffSet / 2 + 1, PicTB.ScaleHeight - 2
        End If
        
        If iFrom = 1 Then r.Left = 0
        
        DrawGradientInRectangle PicTB.hDC, ColorFrom, ColorTo, r, VerticalGradient, False, ColorBorderPic
        PicTB.Refresh
        
    End If
    
    For i = iFrom To iTo
        If TmpTBarItem(i).Visible = True Then
            If ToolbarItem(i).Type = Button Then
                SetRect R1, ToolbarItem(i).Left, ToolbarItem(i).Top, ToolbarItem(i).Width, ToolbarItem(i).R_Height
                DrawBtn PicTB, TmpTBarItem(i).State, R1, i, blnClear
            Else
                DrawSeparator ColorToolbar, ToolbarItem(i).Left, ToolbarItem(i).Top, ToolbarItem(i).R_Height
            End If
        End If
    Next i
    
    PicTB.Picture = PicTB.Image
    
End Sub

'Drawing picTB button
Private Sub DrawBtn(Pic As PictureBox, BtnState As jcBtnState, RBTN As RECT, Index As Long, Optional CaptionRedraw As Boolean = False)
Dim RA As RECT
Dim xIcon As Integer, yIcon As Integer

    Select Case BtnState
        Case STA_PRESSED
                DrawGradientInRectangle Pic.hDC, ColorFromDown, ColorToDown, RBTN, VerticalGradient, True, ColorBorderPic
        Case STA_OVERDOWN
            If ToolbarItem(Index).Style = [Dropdown button] Then
                DrawGradientInRectangle Pic.hDC, BlendColors(ColorTo, vbWhite), BlendColors(ColorFrom, vbWhite), RBTN, VerticalGradient, True, ColorBorderPic
            ElseIf ToolbarItem(Index).Style = [Check button] Then
                DrawGradientInRectangle Pic.hDC, ColorToDown, ColorFromDown, RBTN, VerticalGradient, True, ColorBorderPic
            End If
        Case STA_OVER
                DrawGradientInRectangle Pic.hDC, ColorFromOver, ColorToOver, RBTN, VerticalGradient, True, ColorBorderPic
                TmpTBarItem(Index).Value = False
        Case STA_SELECTED
            If ToolbarItem(Index).Style = [Dropdown button] Then
                DrawGradientInRectangle Pic.hDC, ColorTo, ColorFrom, RBTN, VerticalGradient, True, ColorBorderPic
            ElseIf ToolbarItem(Index).Style = [Check button] Then
                DrawGradientInRectangle Pic.hDC, ColorFromDown, ColorToDown, RBTN, VerticalGradient, True, ColorBorderPic
            End If
        Case STA_NORMAL, STA_DISABLED
            If ToolbarItem(Index).Style = [Dropdown button] Or ToolbarItem(Index).Style = [Check button] Then
                SetRect RA, RBTN.Left, 0, RBTN.Right + 1, Pic.ScaleHeight - 2
                DrawGradientInRectangle Pic.hDC, ColorFrom, ColorTo, RA, VerticalGradient, False, ColorBorderPic
                TmpTBarItem(Index).Value = False
            End If
    End Select
    
    Set_CaptionAndIcon_Rect Index, RBTN, R_Caption, xIcon, yIcon

    If CaptionRedraw Then
    
        If BtnState = STA_DISABLED Then
            DrawCaption Pic, ToolbarItem(Index).Caption, TEXT_INACTIVE, R_Caption, ToolbarItem(Index).font
        Else
            DrawCaption Pic, ToolbarItem(Index).Caption, ToolbarItem(Index).BtnForeColor, R_Caption, ToolbarItem(Index).font
        End If
    
    End If
    
    'Drawing icon picture
    If Not (ToolbarItem(Index).icon Is Nothing) Then
        useMask = ToolbarItem(Index).UseMaskColor
        If BtnState = STA_DISABLED Then
            TransBlt Pic.hDC, xIcon, yIcon, ToolbarItem(Index).Iconsize, ToolbarItem(Index).Iconsize, TmpTBarItem(Index).icon, ToolbarItem(Index).maskColor, , , True, False
        Else
            TransBlt Pic.hDC, xIcon, yIcon, ToolbarItem(Index).Iconsize, ToolbarItem(Index).Iconsize, TmpTBarItem(Index).icon, ToolbarItem(Index).maskColor, , , False, False
        End If
    End If
    
    Pic.Refresh
    
End Sub

Private Sub UpdateCheckValue(Pic As PictureBox, Index As Long)
Dim R1 As RECT

    SetRect R1, ToolbarItem(Index).Left, ToolbarItem(Index).Top, ToolbarItem(Index).Width, ToolbarItem(Index).R_Height
    Pic.Cls
    Pic.Refresh
    DrawBtn Pic, TmpTBarItem(Index).State, R1, Index
    Pic.Refresh
    Pic.Picture = Pic.Image
    
End Sub

Private Sub Set_CaptionAndIcon_Rect(Index As Long, R_Button As RECT, R_Caption As RECT, xIcon As Integer, yIcon As Integer)
    
    'drawing image
    If ToolbarItem(Index).Type <> Button Then Exit Sub
    
    Set PicTB.font = ToolbarItem(Index).font
    If Not (ToolbarItem(Index).icon Is Nothing) Then
        If LenB(ToolbarItem(Index).Caption) <> 0 Then
            Select Case ToolbarItem(Index).BtnAlignment
                Case IconLeftTextRight
                    xIcon = R_Button.Left + m_OffSet
                    yIcon = R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize) / 2
                    'define caption rectangle
                    SetRect R_Caption, R_Button.Left + ToolbarItem(Index).Iconsize + 2 * m_OffSet, R_Button.Top + m_OffSet, R_Button.Left + R_Button.Right - m_OffSet, R_Button.Top + R_Button.Bottom - m_OffSet
                Case IconRightTextLeft
                    xIcon = R_Button.Left + 2 * m_OffSet + PicTB.TextWidth(ToolbarItem(Index).Caption)
                    yIcon = R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize) / 2
                    'define caption rectangle
                    SetRect R_Caption, R_Button.Left + m_OffSet, R_Button.Top + m_OffSet, R_Button.Left + R_Button.Right - 2 * m_OffSet - ToolbarItem(Index).Iconsize, R_Button.Top + R_Button.Bottom - m_OffSet
                Case IconTopTextBottom
                    xIcon = R_Button.Left + (R_Button.Right - ToolbarItem(Index).Iconsize) / 2
                    'yIcon = R_Button.Top + m_OffSet + 1
                    yIcon = R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize - PicTB.TextHeight(ToolbarItem(Index).Caption) - m_OffSet) / 2
                    'define caption rectangle
                    SetRect R_Caption, R_Button.Left + m_OffSet, R_Button.Top + ToolbarItem(Index).Iconsize + 2 * m_OffSet, R_Button.Left + R_Button.Right - m_OffSet, R_Button.Top + R_Button.Bottom - m_OffSet
                Case IconBottomTextTop
                    xIcon = R_Button.Left + (R_Button.Right - ToolbarItem(Index).Iconsize) / 2
                    'yIcon = R_Button.Top + 2 * m_OffSet + PicTB.TextHeight(ToolbarItem(Index).Caption) - 1
                    yIcon = R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize - PicTB.TextHeight(ToolbarItem(Index).Caption) - m_OffSet) / 2 + PicTB.TextHeight(ToolbarItem(Index).Caption) + m_OffSet
                    'define caption rectangle
                    SetRect R_Caption, R_Button.Left + m_OffSet, R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize - PicTB.TextHeight(ToolbarItem(Index).Caption) - m_OffSet) / 2, R_Button.Left + R_Button.Right - m_OffSet, R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize - PicTB.TextHeight(ToolbarItem(Index).Caption) - m_OffSet) / 2 + PicTB.TextHeight(ToolbarItem(Index).Caption)
              End Select
        Else    'only icon
            xIcon = R_Button.Left + (R_Button.Right - ToolbarItem(Index).Iconsize) \ 2
            yIcon = R_Button.Top + (R_Button.Bottom - ToolbarItem(Index).Iconsize) \ 2
        End If
    Else 'no icon
        SetRect R_Caption, R_Button.Left + m_OffSet, R_Button.Top + m_OffSet, R_Button.Left + R_Button.Right - m_OffSet, R_Button.Top + R_Button.Bottom - m_OffSet
    End If
End Sub

Private Sub TransBlt(ByVal DstDC As Long, ByVal DstX As Long, ByVal DstY As Long, ByVal DstW As Long, ByVal DstH As Long, ByVal SrcPic As StdPicture, Optional ByVal TransColor As Long = -1, Optional ByVal BrushColor As Long = -1, Optional ByVal MonoMask As Boolean = False, Optional ByVal isGreyscale As Boolean = False, Optional ByVal XPBlend As Boolean = False)
Dim B As Long, H As Long, f As Long, i As Long, newW As Long
Dim TmpDC As Long, TmpBmp As Long, TmpObj As Long
Dim Sr2DC As Long, Sr2Bmp As Long, Sr2Obj As Long
Dim Data1() As RGBTRIPLE, Data2() As RGBTRIPLE
Dim Info As BITMAPINFO, BrushRGB As RGBTRIPLE, gCol As Long
'Dim hOldOb As Long
Dim SrcDC As Long, tObj As Long ', ttt As Long
    
    If DstW = 0 Or DstH = 0 Then Exit Sub

    SrcDC = CreateCompatibleDC(hDC)

    If DstW < 0 Then DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
    If DstH < 0 Then DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)
    
    If SrcPic.Type = 1 Then 'check if it's an icon or a bitmap
        tObj = SelectObject(SrcDC, SrcPic)
    Else
        Dim hBrush As Long
        tObj = SelectObject(SrcDC, CreateCompatibleBitmap(DstDC, DstW, DstH))
        hBrush = CreateSolidBrush(TransColor) 'MaskColor)
        DrawIconEx SrcDC, 0, 0, SrcPic.Handle, DstW, DstH, 0, hBrush, &H1 Or &H2
        DeleteObject hBrush
    End If

    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    
    ReDim Data1(DstW * DstH * 3 - 1)
    ReDim Data2(UBound(Data1))
    
    With Info.bmiHeader
        .biSize = Len(Info.bmiHeader)
        .biWidth = DstW
        .biHeight = DstH
        .biPlanes = 1
        .biBitCount = 24
    End With

    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0

    If BrushColor > 0 Then
        BrushRGB.rgbBlue = (BrushColor \ &H10000) Mod &H100
        BrushRGB.rgbGreen = (BrushColor \ &H100) Mod &H100
        BrushRGB.rgbRed = BrushColor And &HFF
    End If

    If Not useMask Then TransColor = -1

    newW = DstW - 1

    For H = 0 To DstH - 1
        f = H * DstW
        For B = 0 To newW
            i = f + B
            If GetNearestColor(hDC, CLng(Data2(i).rgbRed) + 256& * Data2(i).rgbGreen + 65536 * Data2(i).rgbBlue) <> TransColor Then
                With Data1(i)
                    If BrushColor > -1 Then
                        If MonoMask Then
                            If (CLng(Data2(i).rgbRed) + Data2(i).rgbGreen + Data2(i).rgbBlue) <= 384 Then Data1(i) = BrushRGB
                        Else
                            Data1(i) = BrushRGB
                        End If
                    Else
                        If isGreyscale Then
                            gCol = CLng(Data2(i).rgbRed * 0.3) + Data2(i).rgbGreen * 0.59 + Data2(i).rgbBlue * 0.11
                            .rgbRed = gCol: .rgbGreen = gCol: .rgbBlue = gCol
                        Else
                            If XPBlend Then
                                .rgbRed = (CLng(.rgbRed) + Data2(i).rgbRed * 2) \ 3
                                .rgbGreen = (CLng(.rgbGreen) + Data2(i).rgbGreen * 2) \ 3
                                .rgbBlue = (CLng(.rgbBlue) + Data2(i).rgbBlue * 2) \ 3
                            Else
                                Data1(i) = Data2(i)
                            End If
                        End If
                    End If
                End With
            End If
        Next B
    Next H

    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0

    Erase Data1, Data2
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
    If SrcPic.Type = 3 Then DeleteObject SelectObject(SrcDC, tObj)
    DeleteDC TmpDC
    DeleteDC Sr2DC
    DeleteObject tObj
    DeleteDC SrcDC
    
End Sub

Private Sub InitialGradToolbar()
Dim r As RECT
    
    PicTB.Cls
    Set PicTB.Picture = Nothing
    SetRect r, 0, 0, PicTB.ScaleWidth, PicTB.ScaleHeight
    DrawVGradientEx PicTB.hDC, ColorTo, ColorFrom, r.Left, r.Top, r.Right, r.Bottom
    APILineEx PicTB.hDC, r.Left, r.Bottom - 1, r.Right, r.Bottom - 1, ColorToolbar
    
    PicTB.Refresh
    PicTB.Picture = PicTB.Image
    
End Sub

Private Function MinimalHeight() As Long
Dim j As Long

    MinimalHeight = 29
    
    For j = 1 To m_ButtonCount
        If ToolbarItem(j).Type = Button Then
            If ToolbarItem(j).Height + 5 > MinimalHeight Then
                MinimalHeight = ToolbarItem(j).Height + 5
            End If
        End If
    Next j
    
    MinimalHeight = MinimalHeight * Screen.TwipsPerPixelY
    
End Function

Private Function fncMakeIcon(frmDC As Long, hBMP As Long, ByVal MaskClr As Long) As Long
    ' where frmDC   (in)  DC of the call window
    '       hBMP    (in)  handle to a bitmap
    '       MaskClr (in)  if = -1 : pixel(0,0)
    ' Return value is a handle to the icon
    '       ipic    (out) icon picture
    
    Dim Bitmapdata As BITMAP  ' bitmap dimension
    Dim iWidth As Long
    Dim iHeight As Long
    Dim SrcDC As Long         ' copy of incoming bitmap
'    Dim hSrc As Long
    Dim oldSrcObj As Long
    Dim MonoDC As Long        ' Mono mask (XOR)
    Dim MonoBmp As Long
    Dim oldMonoObj As Long
    Dim InvertDC As Long      ' Inverted mask (AND)
    Dim InvertBmp As Long
    Dim oldInvertObj As Long
    '
    Dim cBkColor As Long
    Dim icoinfo As ICONINFO

    ' validate input
    If hBMP = 0 Then Exit Function
    
    ' get size of bitmap
    If GetObject(hBMP, Len(Bitmapdata), Bitmapdata) = 0 Then Exit Function
    
    With Bitmapdata
        iWidth = .bmWidth
        iHeight = .bmHeight
    End With
   
    ' create copy of original, we will use it for both masks
    SrcDC = CreateCompatibleDC(0&)
    oldSrcObj = SelectObject(SrcDC, hBMP)
    
    ' get transparecy color
    If MaskClr = -1 Then
        MaskClr = GetPixel(SrcDC, 0, 0)
    End If
   
    ' mono mask (XOR) ............................................
    
    ' create mono DC/Bitmap for mask (XOR mask)
    MonoDC = CreateCompatibleDC(0&)
    MonoBmp = CreateCompatibleBitmap(MonoDC, iWidth, iHeight)
    oldMonoObj = SelectObject(MonoDC, MonoBmp)
    ' Set background of source to the mask color
    cBkColor = GetBkColor(SrcDC)   ' preserve original
    SetBkColor SrcDC, MaskClr
    ' copy bitmap and make monoDC mask in the process
    BitBlt MonoDC, 0, 0, iWidth, iHeight, SrcDC, 0, 0, vbSrcCopy
    ' restore original backcolor
    SetBkColor SrcDC, cBkColor
    ' inverted mask (AND) .................................................

    ' create DC/bitmap for inverted image (AND mask)
    InvertDC = CreateCompatibleDC(frmDC)
    InvertBmp = CreateCompatibleBitmap(frmDC, iWidth, iHeight)
    oldInvertObj = SelectObject(InvertDC, InvertBmp)
    ' copy bitmap into it
    BitBlt InvertDC, 0, 0, iWidth, iHeight, SrcDC, 0, 0, vbSrcCopy
    
    ' Invert background of image to create AND Mask
    SetBkColor InvertDC, vbBlack
    SetTextColor InvertDC, vbWhite
    BitBlt InvertDC, 0, 0, iWidth, iHeight, MonoDC, 0, 0, vbSrcAnd
    
    ' cleanup copy of original
    SelectObject SrcDC, oldSrcObj
    DeleteDC SrcDC
    
    ' Release MonoBmp And InvertBMP
    SelectObject MonoDC, oldMonoObj
    SelectObject InvertDC, oldInvertObj

    With icoinfo
        .fIcon = True
        .xHotspot = 16            ' Doesn't matter here
        .yHotspot = 16
        .hbmMask = MonoBmp
        .hbmColor = InvertBmp
    End With
      
    ' create 'output'
    fncMakeIcon = CreateIconIndirect(icoinfo)
    
CleanUp:
    ' Clean up
    DeleteObject icoinfo.hbmMask
    DeleteObject icoinfo.hbmColor
    DeleteDC MonoDC
    DeleteDC InvertDC
End Function

Private Function fncConvertIconToPic(hIcon As Long) As IPicture
    ' where hIcon   (in)  icon handle
    ' Return value is an interface managing a picture object and its properties
    '          (can be used to set a picture property)

    Dim iGuid As Guid
    Dim pDesc As pictDesc

     '--- check argument
    If hIcon = 0 Then Exit Function
    ' init GUID
    With iGuid
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
    End With
    
    ' fill picture description type
    With pDesc
       .cbSizeofStruct = Len(pDesc)
       .picType = vbPicTypeIcon
       .hImage = hIcon
    End With
    
    OleCreatePictureIndirect pDesc, iGuid, 1, fncConvertIconToPic
    
End Function

Private Sub ConvertToIcon(Index As Integer)
    With ToolbarItem(Index)
        If Not (.icon Is Nothing) Then
            If .icon.Type = vbPicTypeBitmap Then
                Set TmpTBarItem(Index).icon = fncConvertIconToPic(fncMakeIcon(UserControl.hDC, .icon.Handle, -1))
            Else
                Set TmpTBarItem(Index).icon = .icon
            End If
        End If
    End With
End Sub
