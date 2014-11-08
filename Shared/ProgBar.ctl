VERSION 5.00
Begin VB.UserControl ProgBar 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ScaleHeight     =   59
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ToolboxBitmap   =   "ProgBar.ctx":0000
End
Attribute VB_Name = "ProgBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =================================================
' Author:  Kim Pedersen, codemagician@get2net.dk
' Date:    19. May 2000
' Updated: -
' Version: 0.9.1 (First Beta)
'
' Requires:    None
'
' Description:
' A standard implementation of the ProgressBar in
' Common Controls.

' API Declares
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Type PBRANGE
   iLow As Long
   iHigh As Long
End Type

' Progress Class Const
Private Const PROGRESS_CLASS = "msctls_progress32"

' Progress Styles
Private Const PBS_SMOOTH As Long = &H1&
Private Const PBS_VERTICAL As Long = &H4&

' Progress Messages
Private Const PBM_GETRANGE As Long = &H407&
Private Const PBM_SETRANGE As Long = &H401&
Private Const PBM_SETRANGE32 As Long = &H406&
Private Const PBM_SETPOS As Long = &H402&
Private Const PBM_SETSTEP As Long = &H404&
Private Const PBM_STEPIT As Long = &H405&

' Window and other Constants
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000

Public Enum pbOrientation
    [Horizontal] = 0&
    [Vertical] = 1&
End Enum

' Properties
Private m_hWnd As Long
Private m_Max As Long
Private m_Min As Long
Private m_Orientation As pbOrientation
Private m_Smooth As Boolean
Private m_Step As Long
Private m_Value As Long

Private Sub pCreate()
Dim dwStyle As Long
    
    ' Destroy previous control if any
    pDestroy
    
    ' Set styles
    dwStyle = WS_CHILD Or WS_VISIBLE
    
    If m_Orientation = Vertical Then
        dwStyle = dwStyle Or PBS_VERTICAL
    End If
    
    If m_Smooth Then
        dwStyle = dwStyle Or PBS_SMOOTH
    End If
    
    ' Create the progressbar
    m_hWnd = CreateWindowExW(0&, StrPtr(PROGRESS_CLASS), 0&, dwStyle, 0&, 0&, ScaleWidth, ScaleHeight, UserControl.hWnd, 0&, App.hInstance, 0&)
    
    If m_hWnd Then
        ' ProgressBar was created succesfully
        ' Set range and value
        pSetRange
        SendMessageW m_hWnd, PBM_SETPOS, m_Value, 0&
        SendMessageW m_hWnd, PBM_SETSTEP, m_Step, 0&
    End If
    
End Sub

Private Sub pSetRange()
Dim tPR As PBRANGE
Dim tPA As PBRANGE

    ' Try v4.70 PBM_SETRANGE32
    SendMessageW m_hWnd, PBM_SETRANGE32, m_Min, m_Max

    ' Check whether PBM_SETRANGE32 was supported
    tPA.iHigh = SendMessageW(m_hWnd, PBM_GETRANGE, 0&, VarPtr(tPR))
    tPA.iLow = SendMessageW(m_hWnd, PBM_GETRANGE, 1&, VarPtr(tPR))
    
    ' Make sure it worked
    If (tPA.iHigh <> m_Max) Or (tPA.iLow <> m_Min) Then
        ' Use the original set range message otherwhise
        SendMessageW m_hWnd, PBM_SETRANGE, 0&, pMakeDWord(m_Min And &HFFFF&, m_Max And &HFFFF&)
    End If
        
End Sub

Private Sub pDestroy()
    
    ' This sub will destroy any previous
    ' progressbar created
    
    If m_hWnd Then
        ShowWindow m_hWnd, vbHide
        SetParent m_hWnd, 0&
        DestroyWindow m_hWnd
        m_hWnd = 0&
    End If
    
End Sub

Private Function pMakeDWord(lLoWord As Long, lHiWord As Long) As Long
    If (lHiWord And &H8000&) Then
        pMakeDWord = ((lHiWord And &H7FFF&) * &H10000) Or &H80000000 Or lLoWord
    Else
        pMakeDWord = (lHiWord * &H10000) Or lLoWord
    End If
End Function

Private Sub UserControl_Initialize()
    UserControl.ScaleMode = vbPixels
End Sub

Private Sub UserControl_InitProperties()
    m_Min = 0&
    m_Max = 100&
    m_Step = 1&
    m_Smooth = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    ' Read properties
    With PropBag
        m_Min = .ReadProperty("Min", 0&)
        m_Max = .ReadProperty("Max", 100&)
        m_Orientation = .ReadProperty("Orientation", Horizontal)
        m_Smooth = .ReadProperty("Smooth", False)
        m_Step = .ReadProperty("Step", 1&)
    End With
    
    pCreate
    
End Sub

Private Sub UserControl_Resize()
    If m_hWnd Then
        ' Resize Progressbar to fill entire control
        MoveWindow m_hWnd, 0&, 0&, UserControl.ScaleWidth, UserControl.ScaleHeight, 1&
    End If
End Sub

Private Sub UserControl_Terminate()
    ' Destroy the control
    pDestroy
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    ' Save properties
    With PropBag
        .WriteProperty "Min", m_Min, 0&
        .WriteProperty "Max", m_Max, 100&
        .WriteProperty "Orientation", m_Orientation, Horizontal
        .WriteProperty "Smooth", m_Smooth, False
        .WriteProperty "Step", m_Step, 1&
    End With
    
End Sub

Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
Attribute hWnd.VB_MemberFlags = "400"
    hWnd = m_hWnd
End Property

Public Property Get Min() As Long
Attribute Min.VB_Description = "Returns/sets a control's minimum value."
Attribute Min.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Min = m_Min
End Property

Public Property Let Min(ByVal NewValue As Long)
    
    If NewValue <> m_Min Then

        ' Set min value and update new range
        m_Min = NewValue
        
        If m_hWnd Then
            pSetRange
        End If
        
        ' Notify parent object
        PropertyChanged "Min"
    
    End If
    
End Property

Public Property Get Max() As Long
    Max = m_Max
End Property

Public Property Let Max(ByVal NewValue As Long)
Attribute Max.VB_Description = "Returns/sets a control's maximum value."
Attribute Max.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
        
    If NewValue <> m_Max Then
    
        ' Set max value and update the range
        m_Max = NewValue
        
        If m_hWnd Then
            pSetRange
        End If
    
        ' Notify parent object
        PropertyChanged "Max"
        
    End If
    
End Property

Public Property Get Orientation() As pbOrientation
    ' Get current orientation
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal NewValue As pbOrientation)
    
    If NewValue <> m_Orientation Then
        
        m_Orientation = NewValue
        
        If m_hWnd Then
            pCreate
        End If
        
        PropertyChanged "Orientation"
        
    End If
    
End Property

Public Property Get Smooth() As Boolean
   Smooth = m_Smooth
End Property

Public Property Let Smooth(ByVal NewValue As Boolean)

   If NewValue <> m_Smooth Then
      
      m_Smooth = NewValue
      
      If m_hWnd Then
         pCreate
      End If
      
      PropertyChanged "Smooth"
      
   End If
   
End Property

Public Property Get Step() As Long
    Step = m_Step
End Property

Public Property Let Step(ByVal NewValue As Long)
    
    If NewValue <> m_Step Then
        
        m_Step = NewValue
    
        If m_hWnd Then
            SendMessageW m_hWnd, PBM_SETSTEP, m_Step, 0&
        End If
        
        PropertyChanged "Step"
    
    End If
    
End Property

Public Property Get Value() As Long
    Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Long)
Attribute Value.VB_Description = "Returns or sets a control's current Value property."
Attribute Value.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
Attribute Value.VB_MemberFlags = "200"
    
    If NewValue <> m_Value Then
        
        ' Set value
        m_Value = NewValue
        
        If m_hWnd Then
            SendMessageW m_hWnd, PBM_SETPOS, m_Value, 0&
        End If
        
        ' Notify parent object
        PropertyChanged "Value"
        
    End If
    
End Property

Public Sub StepIt()
    
    m_Value = m_Value + m_Step
    
    If m_hWnd Then
        SendMessageW m_hWnd, PBM_STEPIT, 0&, 0&
    End If
    
End Sub

'' ======================================================================================
'' API declares:
'' ======================================================================================
'
'' Memory functions:
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
'    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'' Window functions
'Private Declare Function CreateWindowEX Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
'Private Const WS_VISIBLE = &H10000000
'Private Const WS_CHILD = &H40000000
'Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'Private Const SW_HIDE = 0
'Private Const SW_SHOW = 5
'Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'' Window style bit functions:
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
'    (ByVal hwnd As Long, ByVal nIndex As Long, _
'    ByVal dwNewLong As Long _
'    ) As Long
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
'    (ByVal hwnd As Long, ByVal nIndex As Long _
'    ) As Long
'' Window Long indexes:
'Private Const GWL_EXSTYLE = (-20)
'Private Const GWL_HINSTANCE = (-6)
'Private Const GWL_HWNDPARENT = (-8)
'Private Const GWL_ID = (-12)
'Private Const GWL_STYLE = (-16)
'Private Const GWL_USERDATA = (-21)
'Private Const GWL_WNDPROC = (-4)
'' Style:
'Private Const WS_EX_CLIENTEDGE = &H200&
'Private Const WS_EX_STATICEDGE = &H20000
'
' ' Window relationship functions:
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
'' WIndow position:
'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Const SWP_SHOWWINDOW = &H40
'Private Const SWP_HIDEWINDOW = &H80
'Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
'Private Const SWP_NOACTIVATE = &H10
'Private Const SWP_NOCOPYBITS = &H100
'Private Const SWP_NOMOVE = &H2
'Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
'Private Const SWP_NOREDRAW = &H8
'Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
'Private Const SWP_NOSIZE = &H1
'Private Const SWP_NOZORDER = &H4
'Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
'Private Const HWND_NOTOPMOST = -2
'Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
'' Messages
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'Private Const WM_USER = &H400
'
'Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
'Private Const CLR_INVALID = -1
'
''Style
'Private Const PBS_SMOOTH = &H1
'Private Const PBS_VERTICAL = &H4
'Private Const PBM_SETRANGE = (WM_USER + 1)
'Private Const PBM_SETPOS = (WM_USER + 2)
'Private Const PBM_DELTAPOS = (WM_USER + 3)
'Private Const PBM_SETSTEP = (WM_USER + 4)
'Private Const PBM_STEPIT = (WM_USER + 5)
'Private Const PBM_SETRANGE32 = (WM_USER + 6)
'Private Const PBM_GETRANGE = (WM_USER + 7)
'Private Const PBM_GETPOS = (WM_USER + 8)
'Private Const PBM_SETBARCOLOR = (WM_USER + 9)
'Private Const CCM_FIRST = &H2000
'Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
'Private Const SB_SETBKCOLOR = CCM_SETBKCOLOR
'
'
'' ======================================================================================
'' Implementation:
'' ======================================================================================
'Public Enum EPBBorderStyle
'    epbBorderStyleNone
'    epbBorderStyleSingle
'    epdBorderStyle3d
'End Enum
'
'' ======================================================================================
'' Private variables:
'' ======================================================================================
'Private m_hWnd As Long
'Private m_oBackColor As OLE_COLOR
'Private m_oForeColor As OLE_COLOR
'Private m_bSmooth As Boolean
'Private m_eOrientation As EPBOrientation
'Private m_eBorderStyle As EPBBorderStyle
'Private m_lPosition As Long
'Private m_lMin As Long
'Private m_lMax As Long
'Private m_lStep As Long

'Public Property Get BorderStyle() As EPBBorderStyle
'   BorderStyle = m_eBorderStyle
'End Property
'Property Let BorderStyle(ByVal eBorderStyle As EPBBorderStyle)
'Dim lStyle As Long
'Dim lCStyle As Long
'   If (m_eBorderStyle <> eBorderStyle) Then
'      m_eBorderStyle = eBorderStyle
'      lStyle = GetWindowLong(UserControl.hwnd, GWL_EXSTYLE)
'      If (m_hWnd <> 0) Then
'         lCStyle = GetWindowLong(m_hWnd, GWL_EXSTYLE)
'      End If
'      If (eBorderStyle <> epdBorderStyle3d) Then
'         lStyle = lStyle And Not WS_EX_CLIENTEDGE
'         If (eBorderStyle = epbBorderStyleSingle) Then
'            lCStyle = lCStyle Or WS_EX_STATICEDGE
'         Else
'            lCStyle = lCStyle And Not WS_EX_STATICEDGE
'         End If
'      Else
'         lStyle = lStyle Or WS_EX_CLIENTEDGE
'         lCStyle = lCStyle And Not WS_EX_STATICEDGE
'      End If
'      If (m_hWnd <> 0) Then
'         SetWindowLong m_hWnd, GWL_EXSTYLE, lCStyle
'         SetWindowPos m_hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED
'      End If
'      SetWindowLong UserControl.hwnd, GWL_EXSTYLE, lStyle
'      SetWindowPos UserControl.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_FRAMECHANGED
'      PropertyChanged "BorderStyle"
'   End If
'End Property
'
''Set BackColor
'Public Property Let BackColor(ByVal oNewBackColor As OLE_COLOR)
'   If (oNewBackColor <> m_oBackColor) Then
'      m_oBackColor = oNewBackColor
'      If (m_hWnd <> 0) Then
'         SendMessageLong m_hWnd, SB_SETBKCOLOR, 0, TranslateColor(oNewBackColor)
'      End If
'      PropertyChanged "BackColor"
'   End If
'End Property
'Public Property Get BackColor() As OLE_COLOR
'   BackColor = m_oBackColor
'End Property
'
''SetForeColor
'Public Property Let ForeColor(ByVal oNewForeColor As OLE_COLOR)
'   If (oNewForeColor <> m_oForeColor) Then
'      m_oForeColor = oNewForeColor
'      If (m_hWnd <> 0) Then
'         SendMessageLong m_hWnd, PBM_SETBARCOLOR, 0, TranslateColor(oNewForeColor)
'      End If
'      PropertyChanged "ForeColor"
'   End If
'End Property
'
'Public Property Get ForeColor() As OLE_COLOR
'   ForeColor = m_oForeColor
'End Property
'
'Private Sub pSetRange()

'End Sub
'
'' Convert Automation color to Windows color
'Private Function TranslateColor(ByVal clr As OLE_COLOR, _
'                        Optional hPal As Long = 0) As Long
'    If OleTranslateColor(clr, hPal, TranslateColor) Then
'        TranslateColor = CLR_INVALID
'    End If
'End Function
'
'Private Sub UserControl_InitProperties()
'   Smooth = False
'   Orientation = epbHorizontal
'   pCreate
'   BorderStyle = epbBorderStyleSingle
'   m_oBackColor = UserControl.Ambient.BackColor
'End Sub
'
'Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'   Smooth = PropBag.ReadProperty("Smooth", False)
'   Orientation = PropBag.ReadProperty("Orientation", epbHorizontal)
'   pCreate
'   m_eBorderStyle = -1
'   BorderStyle = PropBag.ReadProperty("BorderStyle", epbBorderStyleSingle)
'   ForeColor = PropBag.ReadProperty("ForeColor", vbHighlight)
'   BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
'   Min = PropBag.ReadProperty("Min", 0)
'   Max = PropBag.ReadProperty("Max", 100)
'   Step = PropBag.ReadProperty("Step", 1)
'End Sub

'Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'   PropBag.WriteProperty "BorderStyle", BorderStyle, epbBorderStyleSingle
'   PropBag.WriteProperty "Smooth", Smooth, False
'   PropBag.WriteProperty "Orientation", Orientation, epbHorizontal
'   PropBag.WriteProperty "ForeColor", m_oForeColor, vbHighlight
'   PropBag.WriteProperty "BackColor", m_oBackColor, vbButtonFace
'   PropBag.WriteProperty "Min", Min, 0
'   PropBag.WriteProperty "Max", Max, 100
'   PropBag.WriteProperty "Step", Step, 1
'End Sub
