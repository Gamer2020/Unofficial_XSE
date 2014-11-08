VERSION 5.00
Begin VB.UserControl xpWellsTab 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00974D37&
   PropertyPages   =   "xpWellsTab.ctx":0000
   ScaleHeight     =   101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   165
   ToolboxBitmap   =   "xpWellsTab.ctx":002A
End
Attribute VB_Name = "xpWellsTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Acknowledgements:

'Ariad Software.
'For letting me look through there ToolBar code
'to see how they use Property Pages

'Manjula Dharmawardhana at www.manjulapra.com.
'For his simple Common Dialog without the .OCX sample

'Special Thanks:
'Steve McMahon ( The Man ) at www.vbaccelerator.com
'for showing us mere mortals how to make quality ActiveX controls.
'Without his generosity and skills, this control would not have happened.

'Planet Source Code, and the people who submit there code:
'For providing the #1 source code site for VB`ers on the net.
    
    Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
'    Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long

'//
'Private m_MouseOver                 As Boolean
'Private lSelectedTab                As Long
Private lHotTab                     As Long
Private MouseInBody                 As Boolean
'Private MouseInTab                  As Boolean
'Private HasFocus                    As Boolean
'Private sAccessKeys                 As String
Private lPrevTab                    As Long
'Property Variables
    Private lTabHeight              As Long
    Private oBackColor              As OLE_COLOR
    Private oForeColor              As OLE_COLOR
    Private oActiveForeColor        As OLE_COLOR
    Private oForeColorHot           As OLE_COLOR
    Private oFrameColor             As OLE_COLOR
    Private oMaskColor              As OLE_COLOR
    Private oTabHighlight1          As OLE_COLOR
    Private oTabHighlight2          As OLE_COLOR
    Private oTabHighlight3          As OLE_COLOR
    Private lTabCount               As Long
    Dim rcTabs()                    As RECT
    Dim rcBody                      As RECT
'Events
Public Event TabPressed(PreviousTab As Long)
Public Event MouseIn(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseOut(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub UserControl_AmbientChanged(PropertyName As String)
    DrawTab
End Sub

Private Sub UserControl_Initialize()
    lTabHeight = 22
End Sub

Private Sub UserControl_InitProperties()
    
    AddTab
    'lTabHeight = 22
    oBackColor = UserControl.Parent.BackColor
    UserControl.BackColor = oBackColor
    oForeColor = vbButtonText
    UserControl.ForeColor = oForeColor
    oActiveForeColor = RGB(56, 80, 152)
    oForeColorHot = RGB(0, 0, 255)
    oFrameColor = RGB(152, 160, 160)
    oMaskColor = RGB(255, 0, 255)
    oTabHighlight1 = RGB(232, 144, 40)
    oTabHighlight2 = RGB(255, 208, 56)
    oTabHighlight3 = RGB(255, 200, 56)
    Set UserControl.font = Ambient.font
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_Show()
    DrawTab
End Sub

Private Sub UserControl_Terminate()
    Erase Tabs
    Erase rcTabs
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim hRgn As Long
Dim i As Long
    
    If Button = vbLeftButton Then
        For i = 1 To lTabCount
            hRgn = CreateRectRgnIndirect(rcTabs(i))
            If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
                lSelectedTab = i
                SelectedTab = i
                DeleteObject hRgn
                Exit For
            Else
                RaiseEvent MouseDown(Button, Shift, X, Y)
            End If
            DeleteObject hRgn
        Next i
    End If
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim hRgn            As Long
Dim i               As Long
Dim lLocalHotTab    As Long
Dim DoRedraw        As Boolean

    If MouseOver(UserControl.hwnd) = True Then
        
        RaiseEvent MouseIn(Button, Shift, X, Y)
        hRgn = CreateRectRgnIndirect(rcBody)
        
        If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
            lHotTab = 0
            If MouseInBody = False Then
                'MouseInTab = False
                MouseInBody = True
                DrawTab
                DeleteObject hRgn
                Exit Sub
            End If
        End If
        
        DeleteObject hRgn
    
        For i = 1 To lTabCount
            hRgn = CreateRectRgnIndirect(rcTabs(i))
            If PtInRegion(hRgn, CLng(X), CLng(Y)) Then
                lLocalHotTab = i
                If lLocalHotTab <> lHotTab Then
                    DoRedraw = True
                    'MouseInTab = True
                    MouseInBody = False
                    lHotTab = i
                    DeleteObject hRgn
                    Exit For
                End If
            End If
            DeleteObject hRgn
        Next i
        
        If DoRedraw = True Then
            DrawTab
        End If
        
        If CLng(X) >= rcTabs(lTabCount).Right Then
            lHotTab = 0
            DrawTab
        End If
        
    Else
        RaiseEvent MouseOut(Button, Shift, X, Y)
        lHotTab = 0
        DrawTab
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        RaiseEvent MouseUp(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_Resize()
    DrawTab
End Sub

'Public Property Get TabHeight() As Long
'    TabHeight = lTabHeight
'End Property

'Public Property Let TabHeight(ByVal NewTabHeight As Long)
'    lTabHeight = NewTabHeight
'    PropertyChanged "TabHeight"
'    DrawTab
'End Property

'Public Property Get BackColor() As OLE_COLOR
'    BackColor = oBackColor
'End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
    oBackColor = NewBackColor
    UserControl.BackColor = oBackColor
    PropertyChanged "BackColor"
    DrawTab
End Property

'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = oForeColor
'End Property
'
'Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
'    oForeColor = NewForeColor
'    UserControl.ForeColor = oForeColor
'    PropertyChanged "ForeColor"
'    DrawTab
'End Property

'Public Property Get ForeColorActive() As OLE_COLOR
'    ForeColorActive = oActiveForeColor
'End Property

Public Property Let ForeColorActive(ByVal NewForeColorActive As OLE_COLOR)
    oActiveForeColor = NewForeColorActive
    PropertyChanged "ForeColorActive"
    DrawTab
End Property

'Public Property Get FrameColor() As OLE_COLOR
'    FrameColor = oFrameColor
'End Property

Public Property Let FrameColor(ByVal NewFrameColor As OLE_COLOR)
    oFrameColor = NewFrameColor
    PropertyChanged "FrameColor"
    DrawTab
End Property

Public Property Get TabWidth(ByVal Index As Long) As Long
    TabWidth = Tabs(Index).TabWidth
End Property

Public Property Let TabWidth(ByVal Index As Long, ByVal NewTabWidth As Long)
    Tabs(Index).TabWidth = NewTabWidth
    DrawTab
    PropertyChanged "TabWidth"
End Property

Public Property Get TabCaption(ByVal Index As Long) As String
    TabCaption = Replace2(Tabs(Index).TabCaption, "&&", "&")
End Property

Public Property Let TabCaption(ByVal Index As Long, ByVal NewTabCaption As String)
    Tabs(Index).TabCaption = Replace2(NewTabCaption, "&", "&&")
'    SetTabAccessKeys
    DrawTab
    PropertyChanged "TabCaption"
End Property

'Public Property Get TabPicture(ByVal Index As Long) As StdPicture
'    Set TabPicture = Tabs(Index).TabIcon
'End Property

'Public Property Set TabPicture(ByVal Index As Long, ByVal NewTabPicture As StdPicture)
'    Set Tabs(Index).TabPicture = NewTabPicture
'    DrawTab
'    PropertyChanged "TabPicture"
'End Property

Public Property Get font() As font
    Set font = UserControl.font
End Property

'Public Property Set font(ByVal NewFont As font)
'    Set UserControl.font = NewFont
'    PropertyChanged "Font"
'    DrawTab
'End Property

'Public Property Get maskColor() As OLE_COLOR
'    maskColor = oMaskColor
'End Property

'Public Property Let maskColor(ByVal NewMaskColor As OLE_COLOR)
'    oMaskColor = NewMaskColor
'    PropertyChanged "MaskColor"
'    DrawTab
'End Property

'Public Property Get TabHighlight1() As OLE_COLOR
'    TabHighlight1 = oTabHighlight1
'End Property

Public Property Let TabHighlight1(ByVal NewTabHighlight1 As OLE_COLOR)
    oTabHighlight1 = NewTabHighlight1
    PropertyChanged "TabHighlight1"
    DrawTab
End Property

'Public Property Get TabHighlight2() As OLE_COLOR
'    TabHighlight2 = oTabHighlight2
'End Property

Public Property Let TabHighlight2(ByVal NewTabHighlight2 As OLE_COLOR)
    oTabHighlight2 = NewTabHighlight2
    PropertyChanged "TabHighlight2"
    DrawTab
End Property

'Public Property Get TabHighlight3() As OLE_COLOR
'    TabHighlight3 = oTabHighlight3
'End Property

Public Property Let TabHighlight3(ByVal NewTabHighlight3 As OLE_COLOR)
    oTabHighlight3 = NewTabHighlight3
    PropertyChanged "TabHighlight3"
    DrawTab
End Property

Public Property Get TabCount() As Long
    TabCount = lTabCount
End Property

'Public Property Get Alignment() As eTabAlignment
'    Alignment = eTab
'End Property
'
'Public Property Let Alignment(ByVal NewAlignment As eTabAlignment)
'    eTab = NewAlignment
'    PropertyChanged "Alignment"
'    DrawTab
'End Property

'Public Property Get ForeColorHot() As OLE_COLOR
'    ForeColorHot = oForeColorHot
'End Property

Public Property Let ForeColorHot(ByVal NewForeColorHot As OLE_COLOR)
    oForeColorHot = NewForeColorHot
    PropertyChanged "ForeColorHot"
End Property

Public Property Get SelectedTab() As Long
    SelectedTab = lSelectedTab
End Property

Public Property Let SelectedTab(ByVal NewSelectedTab As Long)
    lSelectedTab = NewSelectedTab
    RaiseEvent TabPressed(lPrevTab)
    lPrevTab = lSelectedTab
    PropertyChanged "SelectedTab"
    DrawTab
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Long

    With PropBag
'        TabHeight = .ReadProperty("TabHeight", 25)
        BackColor = .ReadProperty("BackColor", vbButtonFace)
        ForeColor = .ReadProperty("ForeColor", vbButtonText)
        ForeColorActive = .ReadProperty("ForeColorActive", RGB(56, 80, 152))
        ForeColorHot = .ReadProperty("ForeColorHot", RGB(0, 0, 255))
        FrameColor = .ReadProperty("FrameColor", RGB(152, 160, 160))
        maskColor = .ReadProperty("MaskColor", RGB(255, 0, 255))
        TabHighlight1 = .ReadProperty("TabHighlight1", RGB(232, 144, 40))
        TabHighlight2 = .ReadProperty("TabHighlight2", RGB(255, 208, 56))
        TabHighlight3 = .ReadProperty("TabHighlight3", RGB(255, 200, 56))
        SelectedTab = .ReadProperty("SelectedTab", 1)
        Set UserControl.font = .ReadProperty("Font", UserControl.font)
        lTabCount = .ReadProperty("TabCount", 0)
    End With
    For i = 1 To lTabCount
        With Tabs(i)
            .TabWidth = PropBag.ReadProperty("TabWidth" & i)
            .TabCaption = PropBag.ReadProperty("TabText" & i)
            'Set .TabIcon = PropBag.ReadProperty("TabPicture" & i)
        End With
    Next i
'    SetTabAccessKeys
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim i As Long
    With PropBag
'        .WriteProperty "Alignment", eTab
        .WriteProperty "TabHeight", lTabHeight
        .WriteProperty "BackColor", oBackColor
        .WriteProperty "ForeColor", oForeColor
        .WriteProperty "ForeColorActive", oActiveForeColor
        .WriteProperty "ForeColorHot", oForeColorHot
        .WriteProperty "FrameColor", oFrameColor
        .WriteProperty "MaskColor", oMaskColor
        .WriteProperty "TabHighlight1", oTabHighlight1
        .WriteProperty "TabHighlight2", oTabHighlight2
        .WriteProperty "TabHighlight3", oTabHighlight3
        .WriteProperty "SelectedTab", lSelectedTab
        .WriteProperty "Font", UserControl.font
        .WriteProperty "TabCount", lTabCount
    End With
    For i = 1 To lTabCount
        With Tabs(i)
            PropBag.WriteProperty "TabWidth" & i, .TabWidth
            PropBag.WriteProperty "TabText" & i, .TabCaption
            'PropBag.WriteProperty "TabPicture" & i, .TabIcon, Nothing
        End With
    Next i
End Sub

Public Sub DrawTab()
Dim RC As RECT
Dim i As Long
Dim rcTemp              As RECT
Dim X                   As Single
Dim Y                   As Single
Dim X1                  As Single
Dim Y1                  As Single
'Dim iOffset             As Long
'Dim iBodyHeight         As Long
Dim pX                  As Single
'Dim pY                  As Single

    RC = GetRect(UserControl.hwnd)
    RC.Top = RC.Top + lTabHeight
    Cls
    DrawASquare UserControl.hDC, RC, oFrameColor
    Y = 2
    X = 2
    'Height of the Tab
    Y1 = lTabHeight
    '//

    'Loop through the Tabs
    For i = 1 To lTabCount
        With Tabs(i)
        'Position the Tabs.
            '.TabLeft = X
            '.TabTop = Y
            X1 = .TabWidth
            '.TabHeight = Y1
        '//
        'Create a RECT area using the above dimentions to draw into.
            With RC
                .Left = X
                .Top = Y
                .Right = .Left + X1
                .Bottom = Y1
                ReDim Preserve rcTabs(i)
                'Save the rect
                rcTabs(i) = RC
                'Left
                DrawALine UserControl.hDC, .Left, .Top + 2, .Left, .Bottom, oFrameColor
                'Top
                DrawALine UserControl.hDC, .Left + 2, .Top, .Right - 1, .Top, oFrameColor
                'Right
                DrawALine UserControl.hDC, .Right, .Top + 2, .Right, .Bottom, oFrameColor
                'Left Corner
                DrawADot UserControl.hDC, .Left + 1, .Top + 1, oFrameColor
                'Right Corner
                DrawADot UserControl.hDC, .Right - 1, .Top + 1, oFrameColor
            End With
            X = (X + 2) + .TabWidth
        End With
    Next i
'Draw the gradients
    For i = 1 To lTabCount
        If i <> lSelectedTab Then
            ClearRect rcTemp
            CopyTheRect rcTemp, rcTabs(i)
            rcTemp.Left = rcTemp.Left + 1
            DrawGradient UserControl.hDC, TranslateColorToRGB(oBackColor, 0, 0, 0, 5), TranslateColorToRGB(oBackColor, 0, 0, 0, -15), rcTemp, lTabHeight - 5
        End If
    Next i

'Draw the selected and hot tab (if required)
    ClearRect RC

    If lTabCount > 0 Then
    
        If lSelectedTab > lTabCount Then
            lSelectedTab = lTabCount
        End If
        
        With Tabs(lSelectedTab)
            rcTabs(lSelectedTab) = ResizeRect(rcTabs(lSelectedTab), 2, 2)
            DrawASquare UserControl.hDC, rcTabs(lSelectedTab), oBackColor, True

            DrawALine UserControl.hDC, rcTabs(lSelectedTab).Left, rcTabs(lSelectedTab).Top + 2, rcTabs(lSelectedTab).Left, rcTabs(lSelectedTab).Bottom - 1, oFrameColor
            DrawALine UserControl.hDC, rcTabs(lSelectedTab).Right, rcTabs(lSelectedTab).Top + 2, rcTabs(lSelectedTab).Right, rcTabs(lSelectedTab).Bottom - 1, oFrameColor
            DrawALine UserControl.hDC, rcTabs(lSelectedTab).Left + 2, rcTabs(lSelectedTab).Top, rcTabs(lSelectedTab).Right - 1, rcTabs(lSelectedTab).Top, oFrameColor
            DrawADot UserControl.hDC, 0, lTabHeight + 1, oFrameColor

            DrawADot UserControl.hDC, rcTabs(lSelectedTab).Left + 1, rcTabs(lSelectedTab).Top + 1, oFrameColor
            DrawADot UserControl.hDC, rcTabs(lSelectedTab).Right - 1, rcTabs(lSelectedTab).Top + 1, oFrameColor

            'HighLights
            'If HasFocus = True Then
                DrawALine UserControl.hDC, rcTabs(lSelectedTab).Left + 2, rcTabs(lSelectedTab).Top, rcTabs(lSelectedTab).Right - 1, rcTabs(lSelectedTab).Top, oTabHighlight1
                DrawALine UserControl.hDC, rcTabs(lSelectedTab).Left + 2, rcTabs(lSelectedTab).Top + 1, rcTabs(lSelectedTab).Right - 1, rcTabs(lSelectedTab).Top + 1, oTabHighlight2
                DrawALine UserControl.hDC, rcTabs(lSelectedTab).Left + 1, rcTabs(lSelectedTab).Top + 2, rcTabs(lSelectedTab).Right, rcTabs(lSelectedTab).Top + 2, oTabHighlight3
            'End If
        End With

        'Hot tab
        If lHotTab <> 0 And lHotTab <> lSelectedTab Then
            With Tabs(lHotTab)
                DrawALine UserControl.hDC, rcTabs(lHotTab).Left + 2, rcTabs(lHotTab).Top, rcTabs(lHotTab).Right - 1, rcTabs(lHotTab).Top, oTabHighlight1
                DrawALine UserControl.hDC, rcTabs(lHotTab).Left + 2, rcTabs(lHotTab).Top + 1, rcTabs(lHotTab).Right - 1, rcTabs(lHotTab).Top + 1, oTabHighlight2
                DrawALine UserControl.hDC, rcTabs(lHotTab).Left + 1, rcTabs(lHotTab).Top + 2, rcTabs(lHotTab).Right, rcTabs(lHotTab).Top + 2, oTabHighlight3
            End With
        End If
    End If
'//
    
    rcBody = GetRect(UserControl.hwnd)
    rcBody.Top = rcBody.Top + lTabHeight
    ClearRect rcTemp
    CopyTheRect rcTemp, rcBody
    ResizeRect rcTemp, -1, -1
    PositionRect rcTemp, 0, 0
    'iBodyHeight = rcTemp.Bottom - rcTemp.Top - 2
'//

'Draw the caption and pictures
    For i = 1 To lTabCount
        ClearRect rcTemp
        CopyTheRect rcTemp, rcTabs(i)
        GetTextRect UserControl.hDC, Tabs(i).TabCaption, Len(Tabs(i).TabCaption), rcTemp
        'GetPictureSize Tabs(i).TabIcon, pX, pY
        PositionRect rcTemp, 8 + pX, ((lTabHeight + rcTemp.Top) - rcTemp.Bottom) \ 2
        'cPic.PaintTransparentPicture UserControl.hdc, Tabs(i).TabIcon, rcTabs(i).Left + 4, ((lTabHeight + rcTemp.Top) - rcTemp.Bottom) / 2, pX, pX, , , oMaskColor
        If i = lSelectedTab Then
            SetTheTextColor UserControl.hDC, oActiveForeColor
            DrawTheText UserControl.hDC, Tabs(i).TabCaption, Len(Tabs(i).TabCaption), rcTemp, Center
        Else
        If i = lHotTab Then
            SetTheTextColor UserControl.hDC, oForeColorHot
            DrawTheText UserControl.hDC, Tabs(i).TabCaption, Len(Tabs(i).TabCaption), rcTemp, Center
        Else
            SetTheTextColor UserControl.hDC, oForeColor
            DrawTheText UserControl.hDC, Tabs(i).TabCaption, Len(Tabs(i).TabCaption), rcTemp, Center
        End If
        End If
    Next i
    'Set cPic = Nothing
'//
Refresh
End Sub

Public Function AddTab(Optional iTabWidth As Long = 50, Optional sTabText As String = vbNullString) As Long ', Optional pTabPicture As StdPicture = Nothing) As Long
    lTabCount = lTabCount + 1
        With Tabs(lTabCount)
            If iTabWidth = 50 And LenB(sTabText) <> 0 Then
                .TabWidth = CalculateTabWidth(sTabText)
            Else
                .TabWidth = iTabWidth
            End If
            If LenB(sTabText) <> 0 Then
                .TabCaption = sTabText
'                SetTabAccessKeys
            Else
                sTabText = CaptionBase & lTabCount
                .TabCaption = sTabText
            End If
            'Set .TabPicture = pTabPicture
        End With
        PropertyChanged "TabCount"
        AddTab = lTabCount
        'DrawTab
End Function

Public Sub RemoveTab()
    If lTabCount > 1 Then
        lTabCount = lTabCount - 1
    End If
    PropertyChanged "TabCount"
End Sub

Public Sub SwapTabs(lTabIndex1 As Long, lTabIndex2 As Long)
Dim tmpCaption As String
Dim tmpWidth As Long
    
    tmpCaption = Tabs(lTabIndex1).TabCaption
    Tabs(lTabIndex1).TabCaption = Tabs(lTabIndex2).TabCaption
    Tabs(lTabIndex2).TabCaption = tmpCaption
    
    tmpWidth = Tabs(lTabIndex1).TabWidth
    Tabs(lTabIndex1).TabWidth = Tabs(lTabIndex2).TabWidth
    Tabs(lTabIndex2).TabWidth = tmpWidth

End Sub
