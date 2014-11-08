VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Copyright © 2009 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hWnd As Long
    uId As Long
    rc As RECT
    hinst As Long
    lpszText As String
    lParam As Long
End Type

Private Const WM_USER As Long = &H400&
Private Const WS_POPUP As Long = &H80000000

Private Const WS_EX_TOPMOST As Long = &H8&
Private Const CW_USEDEFAULT As Long = &H80000000

Private Const HWND_TOPMOST As Long = -1&
Private Const SWP_NOMOVE As Long = &H2&
Private Const SWP_NOACTIVATE As Long = &H10&
Private Const SWP_NOSIZE As Long = &H1&

Private Const TOOLTIPS_CLASS As String = "tooltips_class32"

Private Const TTF_IDISHWND As Long = &H1&
Private Const TTF_SUBCLASS As Long = &H10&

Private Const TTI_ERROR As Long = &H3&
Private Const TTI_ERROR_LARGE As Long = &H6&
Private Const TTI_INFO As Long = &H1&
Private Const TTI_INFO_LARGE As Long = &H4&
Private Const TTI_NONE  As Long = &H0&
Private Const TTI_WARNING As Long = &H2&
Private Const TTI_WARNING_LARGE As Long = &H5&

Private Const TTM_ACTIVATE As Long = &H401&
Private Const TTM_ADDTOOLW As Long = &H432&
Private Const TTM_DELTOOL As Long = &H433&
Private Const TTM_SETMAXTIPWIDTH As Long = &H418&
Private Const TTM_SETTIPBKCOLOR As Long = &H413&
Private Const TTM_SETTIPTEXTCOLOR As Long = &H414&
Private Const TTM_SETTITLEW As Long = &H421&
Private Const TTM_UPDATE As Long = &H41D&
Private Const TTM_UPDATETIPTEXTW As Long = &H439&

Private Const TTS_ALWAYSTIP As Long = &H1&

Public Enum ToolTipIcons
    [NoIcon] = TTI_NONE
    [Information] = TTI_INFO
    [Warning] = TTI_WARNING
    [Error] = TTI_ERROR
    [InformationLarge] = TTI_INFO_LARGE
    [WarningLarge] = TTI_WARNING_LARGE
    [ErrorLarge] = TTI_ERROR_LARGE
End Enum

Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
   
Private m_BackColor As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_hWndOwner As Long
Private m_hWnd As Long
Private m_IsActive As Boolean
Private m_MaxWidth As Long
Private m_Title As String
Private m_ToolTipCount As Long

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    
    ' Make sure it actually changed
    If m_BackColor <> NewValue Then
    
        ' Check if the ToolTip was created
        If m_hWnd <> 0& Then
            
            ' Apply the new BackColor
            m_BackColor = NewValue
            SendMessageW m_hWnd, TTM_SETTIPBKCOLOR, NewValue, 0&
            
        End If
        
    End If
    
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    
    ' Make sure it actually changed
    If m_ForeColor <> NewValue Then
        
        ' Check if the Tooltip was created
        If m_hWnd <> 0& Then
            
            ' Apply the new ForeColor
            m_ForeColor = NewValue
            SendMessageW m_hWnd, TTM_SETTIPTEXTCOLOR, NewValue, 0&
            
        End If
        
    End If
    
End Property

Public Property Get hWndOwner() As Long
    hWndOwner = m_hWndOwner
End Property

Public Property Let hWndOwner(ByVal NewValue As Long)
    
    ' Ensure it's not null
    If NewValue <> 0& Then
    
        ' See if the new owner is different
        ' from the current one
        If NewValue <> m_hWndOwner Then
            
            ' Set the new owner
            m_hWndOwner = NewValue
            
            ' Create the ToolTip
            pCreate
            
        End If
    
    End If
    
End Property

Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Property Get IsActive() As Boolean
    IsActive = m_IsActive
End Property

Public Property Let IsActive(ByVal NewValue As Boolean)
    
    ' Make sure it actually changed
    If NewValue <> m_IsActive Then
        
        ' Check if the Tooltip was created
        If m_hWnd <> 0& Then
        
            ' Activate/deactivate the ToolTip
            SendMessageW m_hWnd, TTM_ACTIVATE, CLng(-CInt(NewValue)), 0&
            
        End If
        
    End If
    
End Property

Public Property Get MaxWidth() As Long
    MaxWidth = m_MaxWidth
End Property

Public Property Let MaxWidth(ByVal NewValue As Long)
    
    ' Make sure it actually changed
    If NewValue <> m_MaxWidth Then
        
        ' Check if the Tooltip was created
        If m_hWnd <> 0& Then
        
            ' Set the MaxWidth
            m_MaxWidth = NewValue
            SendMessageW m_hWnd, TTM_SETMAXTIPWIDTH, 0&, NewValue
            
        End If
        
    End If
    
End Property

Public Property Get ToolTipCount() As Long
    ToolTipCount = m_ToolTipCount
End Property

Public Sub AddToolTip(ByVal hWndControl As Long, ByRef Text As String)
Dim ti As TOOLINFO
Dim rc As RECT
    
    ' Ensure the ToolTip was created
    If m_hWnd <> 0& Then
    
        ' Ensure it's a valid handle
        If hWndControl <> 0& Then
        
            ' Avoid overflow
            If m_ToolTipCount < 2147483647 Then
            
                GetClientRect hWndControl, rc
            
                ' Fill the TOOLINFO structure
                With ti
                    .cbSize = Len(ti)
                    .uFlags = TTF_IDISHWND Or TTF_SUBCLASS
                    .hWnd = hWndControl
                    .hinst = App.hInstance
                    .uId = hWndControl
                    .lpszText = Text
                    .rc = rc
                End With
                
                ' Add the new tool
                SendMessageW m_hWnd, TTM_ADDTOOLW, 0&, VarPtr(ti)
                
                ' Increment counter
                m_ToolTipCount = m_ToolTipCount + 1&
            
            End If
            
        End If
        
    End If

End Sub

Public Sub Refresh()
    
    ' Ensure the ToolTip was created
    If hWnd <> 0& Then
        SendMessageW m_hWnd, TTM_UPDATE, 0&, 0&
    End If
    
End Sub

Public Sub RemoveToolTip(ByVal hWndControl As Long)
Dim ti As TOOLINFO
    
    ' Ensure the ToolTip was created
    If m_hWnd <> 0& Then
        
        ' Ensure it's a valid handle
        If hWndControl <> 0& Then
            
            ' Check if there's a list one tool
            If m_ToolTipCount Then
            
                ' Fill the TOOLINFO structure
                With ti
                   .cbSize = Len(ti)
                   .hWnd = hWndControl
                   .uId = hWndControl
                End With
                
                ' Remove the tool
                SendMessageW m_hWnd, TTM_DELTOOL, 0&, VarPtr(ti)
                
                ' Decrement the counter
                m_ToolTipCount = m_ToolTipCount - 1&
            
            End If
            
        End If
    
    End If

End Sub

Public Sub SetTitle(ByRef Title As String, Optional ByVal Icon As ToolTipIcons)
    
    ' Check if the ToolTip was created
    If m_hWnd <> 0& Then
        SendMessageW m_hWnd, TTM_SETTITLEW, Icon, StrPtr(Title)
    End If

End Sub

Public Sub UpdateText(ByVal hWndControl As Long, ByRef NewText As String)
Dim ti As TOOLINFO
   
    ' Ensure the ToolTip was created
    If m_hWnd <> 0& Then
       
       ' Ensure it's a valid handle
       If hWndControl <> 0& Then
       
            ' Fill the TOOLINFO structure
            With ti
               .cbSize = Len(ti)
               .hWnd = hWndControl
               .uId = hWndControl
               .lpszText = NewText
            End With
            
            ' Update the text
            SendMessageW m_hWnd, TTM_UPDATETIPTEXTW, 0&, VarPtr(ti)
       
       End If
       
    End If

End Sub

Private Sub pCreate()

    ' Remove the previous ToolTip, if any
    pDestroy

    ' Create a new ToolTip
    m_hWnd = CreateWindowExW(WS_EX_TOPMOST, StrPtr(TOOLTIPS_CLASS), 0&, WS_POPUP Or TTS_ALWAYSTIP, _
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, m_hWndOwner, 0&, App.hInstance, ByVal 0&)
    
    ' If the ToolTip was created successfully
    If m_hWnd <> 0& Then
        
        ' Set its position
        SetWindowPos m_hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
        
    End If

End Sub

Private Sub pDestroy()
    
    ' Check if there a Tooltip
    If m_hWnd <> 0& Then
        
        ' Destroy it
        DestroyWindow m_hWnd
        m_hWnd = 0&
        
    End If
    
End Sub

Private Sub Class_Initialize()
    m_BackColor = vbInfoBackground
    m_ForeColor = vbInfoText
    m_IsActive = True
End Sub

Private Sub Class_Terminate()
    pDestroy
End Sub
