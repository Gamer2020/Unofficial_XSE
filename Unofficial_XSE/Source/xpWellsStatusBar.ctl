VERSION 5.00
Begin VB.UserControl xpWellsStatusBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5940
   ControlContainer=   -1  'True
   PropertyPages   =   "xpWellsStatusBar.ctx":0000
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   ToolboxBitmap   =   "xpWellsStatusBar.ctx":0016
End
Attribute VB_Name = "xpWellsStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Control Name:          xpWellsStatusBar
'Created:               01/02/2003
'Author:                Richard Wells.

'Acknowledgements:

'Ariad Software.
'For letting me look through there ToolBar code
'to see how they use Property Pages

'Manjula Dharmawardhana at www.manjulapra.com
'For his simple Common Dialog without the .OCX sample

'Special Thanks:
'Steve McMahon ( The Man ) at www.vbaccelerator.com
'for showing us mere mortals how to make quality ActiveX controls.
'Without his generosity and skills, this control would not have happened.


'API Stuff.
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function ReleaseCapture Lib "user32" () As Long

    'GDI and regions.
        Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
        Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal Y As Long) As Long
    '//
    
    'System Color Stuff
    Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long
    Private Const CLR_INVALID = -1
    
    'Text Stuff
    Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
    Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
    Private Const DT_CALCRECT = &H400
    Private Const DT_WORDBREAK = &H10
    Private Const DT_CENTER = &H1 Or DT_WORDBREAK Or &H4
    Private Const DT_WORD_ELLIPSIS = &H40000
    
    'Graphics Stuff
    Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
    Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
    Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
    
    Private Const PS_SOLID = 0
    
    'Reigons and Rects
    Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
    Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
    Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal Y As Long) As Long
    Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, lpSourceRect As RECT) As Long
    Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
    
    Public Enum DrawTextFlags
        [Word Break] = DT_WORDBREAK
        Center = DT_CENTER
        [Use Ellipsis] = DT_WORD_ELLIPSIS
    End Enum

'Gripper Stuff
    Private Const WM_NCLBUTTONDOWN = &HA1
    Private Const HTBOTTOMRIGHT = 17
    Private rcGripper               As RECT
'//
    
'Panel Stuff.
    Private lPanelCount             As Integer
    Private rcPanel()               As RECT
    
    'Used for Click and DblClick Events
    Private PanelNum                As Long
    '//
'//

'Colors
    'Panel colors and global mask color.
        Private oBackColor          As OLE_COLOR
        Private oForeColor          As OLE_COLOR
        Private oMaskColor          As OLE_COLOR
        Private oDissColor          As OLE_COLOR
    '//
'//

Event MouseDownInPanel(lPanel As Long)
Event DblClick(lPanelNumber As Long)

Private Function TranslateColorToRGB(ByVal oClr As OLE_COLOR, ByRef r As Long, ByRef G As Long, ByRef B As Long, Optional iOffset As Long = 0, Optional hPal As Long = 0) As OLE_COLOR
Dim iRGB As Long
    
    If OleTranslateColor(oClr, hPal, iRGB) Then
        TranslateColorToRGB = CLR_INVALID
    End If
    
    r = ((iRGB And &HFF&) + iOffset)
    G = (((iRGB And &HFF00&) \ &H100) + iOffset)
    B = (((iRGB And &HFF0000) \ &H10000) + iOffset)
    
    If r < 0 Then
        r = 0
    Else
        If r > 255 Then
            r = 255
        End If
    End If

    If G < 0 Then
        G = 0
    Else
        If G > 255 Then
            G = 255
        End If
    End If

    If B < 0 Then
        B = 0
    Else
        If B > 255 Then
            B = 255
        End If
    End If
    
    TranslateColorToRGB = RGB(r, G, B)
    
End Function

Private Sub DrawASquare(DestDC As Long, RC As RECT, oColor As OLE_COLOR, Optional bFillRect As Boolean)
Dim iBrush As Long
    
    oColor = TranslateColorToRGB(oColor, 0, 0, 0)
    iBrush = CreateSolidBrush(oColor)
    
    If bFillRect = True Then
        FillRect DestDC, RC, iBrush
    Else
        FrameRect DestDC, RC, iBrush
    End If
    
    DeleteObject iBrush
    
End Sub

Private Sub DrawALine(DestDC As Long, x As Long, Y As Long, X1 As Long, Y1 As Long, oColor As OLE_COLOR, Optional iWidth As Long = 1)
Dim pt As POINTAPI
Dim iPen As Long
Dim iPen1 As Long
    
    iPen = CreatePen(PS_SOLID, iWidth, oColor)
    iPen1 = SelectObject(DestDC, iPen)
    
    MoveToEx DestDC, x, Y, pt
    LineTo DestDC, X1, Y1
    
    SelectObject DestDC, iPen1
    DeleteObject iPen1
    DeleteObject iPen
    
End Sub

Private Function GetRect(iHwnd As Long) As RECT
    GetClientRect iHwnd, GetRect
End Function

Private Sub SetTheTextColor(DestDC As Long, oColor As OLE_COLOR)
    SetTextColor DestDC, oColor
End Sub

Private Sub DrawTheText(DestDC As Long, sText As String, iTextLength As Long, RC As RECT, DTF As DrawTextFlags)
    DrawTextEx DestDC, sText, iTextLength, RC, DTF, ByVal 0&
End Sub

Private Sub GetTextRect(DestDC As Long, sText As String, iTextLength As Long, RC As RECT)
    DrawTextEx DestDC, sText, iTextLength, RC, DT_CALCRECT Or DT_WORDBREAK, ByVal 0&
End Sub

Private Sub CopyTheRect(DestinationRECT As RECT, SourceRECT As RECT)
    CopyRect DestinationRECT, SourceRECT
End Sub

Private Sub PositionRect(RC As RECT, ByVal x As Long, ByVal Y As Long)
    OffsetRect RC, x, Y
End Sub

Private Function ResizeRect(RC As RECT, X1 As Long, Y1 As Long) As RECT
    InflateRect RC, X1, Y1
    ResizeRect = RC
End Function

Private Sub UserControl_AmbientChanged(PropertyName As String)
    DrawStatusBar
End Sub

Private Sub UserControl_DblClick()
    'If Panels(PanelNum).pEnabled = True Then
        RaiseEvent DblClick(PanelNum)
    'End If
End Sub

Private Sub UserControl_InitProperties()
    oBackColor = vbButtonFace
    oForeColor = vbButtonText
    oDissColor = vbGrayText
    oMaskColor = RGB(255, 0, 255)
End Sub

Private Sub UserControl_Show()
    DrawStatusBar
End Sub

Private Sub UserControl_Terminate()
    Erase rcPanel
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
'Dim pt As POINTAPI
Dim hRgn As Long
Dim i As Long
    
    If Button = vbLeftButton Then
    
        PanelNum = 0
    
        hRgn = CreateRectRgnIndirect(rcGripper)
        
        If PtInRegion(hRgn, CLng(x), CLng(Y)) Then
            SizeByGripper frmMain.hWnd
            DeleteObject hRgn
            Exit Sub
        Else
            DeleteObject hRgn
        End If
    
        For i = 1 To lPanelCount
            hRgn = CreateRectRgnIndirect(rcPanel(i))
            If PtInRegion(hRgn, CLng(x), CLng(Y)) Then
                'If Panels(i).pEnabled = True Then
                    PanelNum = i
                    RaiseEvent MouseDownInPanel(i)
                'End If
            End If
            DeleteObject hRgn
        Next i
    
    End If
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim hRgn As Long
Dim i As Long
    
    hRgn = CreateRectRgnIndirect(rcGripper)
    
    If PtInRegion(hRgn, CLng(x), CLng(Y)) Then
        UserControl.MousePointer = vbSizeNWSE
        DeleteObject hRgn
        Exit Sub
    Else
        UserControl.MousePointer = vbDefault
        DeleteObject hRgn
    End If
    
    For i = 1 To lPanelCount
        hRgn = CreateRectRgnIndirect(rcPanel(i))
        If PtInRegion(hRgn, CLng(x), CLng(Y)) Then
            Extender.ToolTipText = Panels(i).ToolTipTxt
        End If
        DeleteObject hRgn
    Next i
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ReleaseCapture
End Sub

'Public Property Get BackColor() As OLE_COLOR
'    BackColor = oBackColor
'End Property

Public Property Let BackColor(ByVal NewBackColor As OLE_COLOR)
    oBackColor = NewBackColor
    UserControl.BackColor = oBackColor
    DrawStatusBar
    PropertyChanged "BackColor"
End Property

'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = oForeColor
'End Property

'Public Property Let ForeColor(ByVal NewForeColor As OLE_COLOR)
'    oForeColor = NewForeColor
'    PropertyChanged "ForeColor"
'    DrawStatusBar False
'End Property

'Public Property Get NumberOfPanels() As Long
'    NumberOfPanels = lPanelCount
'End Property

Public Property Get PanelWidth(ByVal Index As Long) As Long
    PanelWidth = Panels(Index).ClientWidth
End Property

Public Property Let PanelWidth(ByVal Index As Long, ByVal PanelWidth As Long)
    Panels(Index).ClientWidth = PanelWidth
    DrawStatusBar
    PropertyChanged "PWidth"
End Property

'Public Property Get PanelCaption(ByVal Index As Long) As String
'    PanelCaption = Panels(Index).PanelText
'End Property

Public Property Let PanelCaption(ByVal Index As Long, ByVal NewPanelCaption As String)
    Panels(Index).PanelText = NewPanelCaption
    DrawStatusBar False
    PropertyChanged "pText"
End Property

'Public Property Get ToolTipText(ByVal Index As Long) As String
'    ToolTipText = Panels(Index).ToolTipTxt
'End Property

'Public Property Let ToolTipText(ByVal Index As Long, ByVal NewToolTipText As String)
'    Panels(Index).ToolTipTxt = NewToolTipText
'    PropertyChanged "pTTText"
'End Property

Public Property Get PanelEnabled(ByVal Index As Long) As Boolean
    PanelEnabled = Panels(Index).pEnabled
End Property

Public Property Let PanelEnabled(ByVal Index As Long, ByVal NewEnabled As Boolean)
    Panels(Index).pEnabled = NewEnabled
    DrawStatusBar False
    PropertyChanged "pEnabled"
End Property

'Public Property Get maskColor() As OLE_COLOR
'    maskColor = oMaskColor
'End Property

'Public Property Let maskColor(ByVal NewMaskColor As OLE_COLOR)
'    oMaskColor = NewMaskColor
'    PropertyChanged "MaskColor"
'    DrawStatusBar False
'End Property

'Public Property Get font() As font
'    Set font = UserControl.font
'End Property

'Public Property Set font(ByVal NewFont As font)
'    Set UserControl.font = NewFont
'    PropertyChanged "Font"
'    DrawStatusBar False
'End Property

'Public Property Get ForeColorDissabled() As OLE_COLOR
'    ForeColorDissabled = oDissColor
'End Property

Public Property Let ForeColorDissabled(ByVal NewDissColor As OLE_COLOR)
    oDissColor = NewDissColor
    PropertyChanged "ForeColorDissabled"
    DrawStatusBar False
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim i As Long
On Error GoTo ERH:
    With PropBag
        BackColor = .ReadProperty("BackColor", vbButtonFace)
        ForeColor = .ReadProperty("ForeColor", vbButtonText)
        ForeColorDissabled = .ReadProperty("ForeColorDissabled", vbGrayText)
        Set UserControl.font = .ReadProperty("Font", UserControl.Ambient.font)
        lPanelCount = .ReadProperty("NumberOfPanels", 0)
        maskColor = .ReadProperty("MaskColor", RGB(255, 0, 255))
    End With
    For i = 1 To lPanelCount
        With Panels(i)
            .ClientWidth = PropBag.ReadProperty("PWidth" & i)
            .ToolTipTxt = PropBag.ReadProperty("pTTText" & i)
            .PanelText = PropBag.ReadProperty("pText" & i)
            .pEnabled = PropBag.ReadProperty("pEnabled" & i)
        End With
    Next i
Exit Sub
ERH:
If Err.Number = 327 Then
    Err.Clear
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim i As Long
    With PropBag
        .WriteProperty "BackColor", oBackColor
        .WriteProperty "ForeColor", oForeColor
        .WriteProperty "ForeColorDissabled", oDissColor
        .WriteProperty "Font", UserControl.font
        .WriteProperty "NumberOfPanels", lPanelCount
        .WriteProperty "MaskColor", oMaskColor
    End With

    For i = 1 To lPanelCount
        With Panels(i)
            PropBag.WriteProperty "PWidth" & i, .ClientWidth
            PropBag.WriteProperty "pText" & i, .PanelText
            PropBag.WriteProperty "pTTText" & i, .ToolTipTxt
            PropBag.WriteProperty "pEnabled" & i, .pEnabled
        End With
    Next i
    
End Sub

Private Sub UserControl_Resize()
    DrawStatusBar
End Sub

Private Sub DrawGripper()

    With rcGripper
        .Left = UserControl.ScaleWidth - 15
        .Right = UserControl.ScaleWidth
        .Bottom = UserControl.ScaleHeight
        .Top = UserControl.ScaleHeight - 15
    End With
    
    With UserControl
        'Retain the area
        DrawASquare .hDC, rcGripper, .BackColor, True
        DrawALine .hDC, rcGripper.Left, rcGripper.Bottom - 1, rcGripper.Right, rcGripper.Bottom - 1, TranslateColorToRGB(oBackColor, 0, 0, 0, -15), 2
        DrawALine .hDC, rcGripper.Left, rcGripper.Bottom - 3, rcGripper.Right, rcGripper.Bottom - 3, TranslateColorToRGB(oBackColor, 0, 0, 0, -8), 2
        
        DrawALine .hDC, .ScaleWidth - 3, .ScaleHeight - 3, .ScaleWidth - 3, .ScaleHeight - 3, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
        DrawALine .hDC, .ScaleWidth - 7, .ScaleHeight - 3, .ScaleWidth - 7, .ScaleHeight - 3, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
        DrawALine .hDC, .ScaleWidth - 11, .ScaleHeight - 3, .ScaleWidth - 11, .ScaleHeight - 3, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
    
        DrawALine .hDC, .ScaleWidth - 3, .ScaleHeight - 7, .ScaleWidth - 3, .ScaleHeight - 7, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
        DrawALine .hDC, .ScaleWidth - 7, .ScaleHeight - 7, .ScaleWidth - 7, .ScaleHeight - 7, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
    
        DrawALine .hDC, .ScaleWidth - 3, .ScaleHeight - 11, .ScaleWidth - 3, .ScaleHeight - 11, TranslateColorToRGB(.BackColor, 0, 0, 0, 50), 2
    
        DrawALine .hDC, .ScaleWidth - 4, .ScaleHeight - 4, .ScaleWidth - 4, .ScaleHeight - 4, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
        DrawALine .hDC, .ScaleWidth - 8, .ScaleHeight - 4, .ScaleWidth - 8, .ScaleHeight - 4, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
        DrawALine .hDC, .ScaleWidth - 12, .ScaleHeight - 4, .ScaleWidth - 12, .ScaleHeight - 4, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
    
        DrawALine .hDC, .ScaleWidth - 4, .ScaleHeight - 8, .ScaleWidth - 4, .ScaleHeight - 8, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
        DrawALine .hDC, .ScaleWidth - 8, .ScaleHeight - 8, .ScaleWidth - 8, .ScaleHeight - 8, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
    
        DrawALine .hDC, .ScaleWidth - 4, .ScaleHeight - 12, .ScaleWidth - 4, .ScaleHeight - 12, TranslateColorToRGB(.BackColor, 0, 0, 0, -50), 2
        UserControl.Refresh
    End With

End Sub

'Public Function AddPanel(Optional lPanelWidth As Long = 100, Optional sPanelText As String = vbNullString, Optional sToolTip As String = vbNullString, Optional bEnabled As Boolean = True) As Long ', Optional pPanelPicture As StdPicture = Nothing) As Long
'    lPanelCount = lPanelCount + 1
'        With Panels(lPanelCount)
'            .ClientWidth = lPanelWidth
'            .ToolTipTxt = sToolTip
'            .PanelText = sPanelText
'            .pEnabled = bEnabled
'            'Set .PanelPicture = pPanelPicture
'        End With
'        PropertyChanged "NumberOfPanels"
'        AddPanel = lPanelCount
'        DrawStatusBar
'End Function
'
'Public Sub DeletePanel()
'    If lPanelCount > 1 Then
'        lPanelCount = lPanelCount - 1
'    End If
'    PropertyChanged "NumberOfPanels"
'    DrawStatusBar
'End Sub

Private Sub DrawStatusBar(Optional FullRedraw As Boolean = True)
Dim i                   As Long
Dim RC                  As RECT
Dim rcTemp              As RECT
Dim x                   As Long
Dim Y                   As Long
Dim X1                  As Long
Dim Y1                  As Long
Dim iOffset             As Long
Dim pX                  As Long
'Dim pY                  As Long
iOffset = 36

If FullRedraw = True Then
With UserControl
    'Control Shading Lines.
    Cls
    'Top lines
    DrawALine .hDC, 0, 0, .ScaleWidth, 0, TranslateColorToRGB(oBackColor, 0, 0, 0, -45)
    For i = 1 To 4
        DrawALine .hDC, 0, i, .ScaleWidth, i, TranslateColorToRGB(oBackColor, 0, 0, 0, iOffset)
        iOffset = iOffset - 9
    Next i
    '//
    
    'Bottom Lines
    DrawALine .hDC, 0, .ScaleHeight - 1, .ScaleWidth, .ScaleHeight - 1, TranslateColorToRGB(oBackColor, 0, 0, 0, -15), 2
    DrawALine .hDC, 0, .ScaleHeight - 3, .ScaleWidth, .ScaleHeight - 3, TranslateColorToRGB(oBackColor, 0, 0, 0, -8), 2
    '//
'//
End With
End If
'The Panels.
    '******************* Dimentions. **********************
    'X = Left of the panel
    'Y = Top of the panel
    'X1 = Width of the panel
    'Y1 = Height of the panel
    '******************************************************
    
    'Start the panel 5 pixels down from the top edge.
    Y = 5
    '//
    'Height of the panel
    Y1 = UserControl.ScaleHeight - 4
    '//
    
    'Loop through the panels
    For i = 1 To lPanelCount
        With Panels(i)
        'Position the panel.
            '.ClientLeft = X
            '.ClientTop = Y
            'X1 is taken from property "PanelWidth"
            X1 = .ClientWidth
            '//
            '.ClientHeight = Y1
        '//
        'Create a RECT area using the above dimentions to draw into.
            With RC
                .Left = x
                .Top = Y
                .Right = .Left + X1
                .Bottom = Y1
            End With
            ReDim Preserve rcPanel(i)
            rcPanel(i) = RC
            ResizeRect rcPanel(i), -2, 0
        '//
        
        If FullRedraw = True Then
        'Draw the seperators taking into acount the first and last
        'panel seperators are different.
            If i <> 1 Then
            'This will draw the left line ( The lighter shade )
            'so the first panel does not need one
                DrawALine UserControl.hDC, x, Y, x, Y1, TranslateColorToRGB(oBackColor, 0, 0, 0, 50)
            '//
            End If
            If i <> lPanelCount Then
            'This will draw the right line ( The darker shade )
            'Every panel will have this line exept the last
            'panel has this line positioned differently.
                DrawALine UserControl.hDC, RC.Right - 1, Y, RC.Right - 1, Y1, TranslateColorToRGB(oBackColor, 0, 0, 0, -50)
            '//
            Else
                If i = lPanelCount Then
                'Lines for the last panel.
                    DrawALine UserControl.hDC, RC.Right - 1, Y, RC.Right - 1, Y1, TranslateColorToRGB(oBackColor, 0, 0, 0, 50)
                    DrawALine UserControl.hDC, RC.Right - 2, Y, RC.Right - 2, Y1, TranslateColorToRGB(oBackColor, 0, 0, 0, -50)
                '//
                End If
            End If
        '//
        End If
        
        DrawASquare UserControl.hDC, rcPanel(i), oBackColor, True

        'Create a temporary RECT to draw some text into.
            rcTemp = GetRect(UserControl.hWnd)
            GetTextRect UserControl.hDC, .PanelText, Len(.PanelText), rcTemp
        '//
        'Copy the temporary RECT
            CopyTheRect RC, rcTemp
        '//
        'Position our RECT
            RC.Left = x
            RC.Right = ((RC.Left + X1) - 6) - pX
        '//
        'Draw the text into our new panel.
            
            If .pEnabled = True Then
                SetTheTextColor UserControl.hDC, oForeColor
            Else
                SetTheTextColor UserControl.hDC, oDissColor
            End If
            
            PositionRect RC, 2 + pX + 4, (ScaleHeight - RC.Bottom) \ 2
            DrawTheText UserControl.hDC, .PanelText, Len(.PanelText), RC, [Use Ellipsis]

        'Dont forget to move the X ( Or left )
        'for the next panel.
            x = x + .ClientWidth
        '//
        End With
    Next i
'//
    DrawGripper

End Sub

Private Sub SizeByGripper(ByVal iHwnd As Long)
  ReleaseCapture
  SendMessage iHwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
End Sub
