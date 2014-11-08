VERSION 5.00
Begin VB.UserControl vcProgress 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ScaleHeight     =   59
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   101
   ToolboxBitmap   =   "vcProgress.ctx":0000
End
Attribute VB_Name = "vcProgress"
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
'
' -------------------------------------------------
' Visit vbCode Magician, free source for VB
' programmers. http://hjem.get2net.dk/vcoders/cm
' =================================================

' Properties
Private m_hWnd As Long
Private m_Max As Long
Private m_Min As Long
Private m_Orientation As OrientationConst
Private m_Value As Long

' API Declares
Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' Progress Class Const
Private Const PROGRESS_CLASS = "msctls_progress32"

' Progress Styles
Private Const PBS_SMOOTH = &H1&
Private Const PBS_VERTICAL = &H4&

' Common Controls shared constants
Private Const WM_USER = &H400&

' Progress Messages
Private Const PBM_SETRANGE = (WM_USER + 1)
Private Const PBM_SETPOS = (WM_USER + 2)

' Window and other Constants
Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000

Public Enum OrientationConst
    vcHorizontal = 0
    vcVertical = 1
End Enum

Private Sub CreateProgressBar()
Dim dwStyle As Long
    
    ' Destroy previous control if any
    DestroyProgressBar
    
    ' Set styles
    dwStyle = WS_CHILD Or WS_VISIBLE Or PBS_SMOOTH
    
    If m_Orientation = vcVertical Then
        dwStyle = dwStyle Or PBS_VERTICAL
    End If
    
    ' Create the progressbar
    m_hWnd = CreateWindowExW(0&, StrPtr(PROGRESS_CLASS), 0&, dwStyle, 0&, 0&, ScaleWidth, ScaleHeight, UserControl.hWnd, 0&, App.hInstance, 0&)
    
    If m_hWnd Then
        ' ProgressBar was created succesfully
        ' Set range and value
        SendMessageW m_hWnd, PBM_SETRANGE, 0&, MakeDWord(m_Min, m_Max)
        SendMessageW m_hWnd, PBM_SETPOS, m_Value, 0&
    End If
    
End Sub

Private Sub DestroyProgressBar()
    
    ' This sub will destroy any previous
    ' progressbar created
    
    If m_hWnd Then
        ' A progressbar already exist. This one
        ' will be destroyed.
        DestroyWindow m_hWnd
        m_hWnd = 0&
    End If
    
End Sub

Private Function MakeDWord(lLoWord As Long, lHiWord As Long) As Long
    If (lHiWord And &H8000&) Then
        MakeDWord = ((lHiWord And &H7FFF&) * &H10000) Or &H80000000 Or lLoWord
    Else
        MakeDWord = (lHiWord * &H10000) Or lLoWord
    End If
End Function

Private Sub UserControl_InitProperties()
    
    ' Set initial properties
    m_Max = 100&
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    ' Read properties
    With PropBag
        m_Min = .ReadProperty("Min", 0&)
        m_Max = .ReadProperty("Max", 100&)
        m_Orientation = .ReadProperty("Orientation", vcHorizontal)
        m_Value = .ReadProperty("Value", 0&)
    End With
    
End Sub

Private Sub UserControl_Resize()
    If m_hWnd Then
        ' Resize Progressbar to fill entire control
        MoveWindow m_hWnd, 0&, 0&, ScaleWidth, ScaleHeight, 1&
    End If
End Sub

Private Sub UserControl_Show()
    ' Create Progressbar
    If m_hWnd = 0& Then
        CreateProgressBar
    End If
End Sub

Private Sub UserControl_Terminate()
    ' Destroy the control
    DestroyProgressBar
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    ' Save properties
    With PropBag
        .WriteProperty "Min", m_Min, 0&
        .WriteProperty "Max", m_Max, 100&
        .WriteProperty "Orientation", m_Orientation, vcHorizontal
        .WriteProperty "Value", m_Value, 0&
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
    ' Get max value
    Min = m_Min
End Property

Public Property Let Min(ByVal NewValue As Long)
    
    If NewValue >= 0& Then
        
        If NewValue < 65536 Then
        
            If NewValue <= m_Max Then
        
                If NewValue <> m_Min Then
            
                    ' Set min value and update new range
                    m_Min = NewValue
                    
                    If m_hWnd Then
                        SendMessageW m_hWnd, PBM_SETRANGE, 0&, MakeDWord(m_Min, m_Max)
                    End If
                    
                    ' Notify parent object
                    PropertyChanged "Min"
                
                End If
            
            End If
        
        End If
    
    End If
    
End Property

Public Property Get Max() As Long
    ' Get max value
    Max = m_Max
End Property

Public Property Let Max(ByVal NewValue As Long)
Attribute Max.VB_Description = "Returns/sets a control's maximum value."
Attribute Max.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
    
    If NewValue >= 0& Then
    
        If NewValue < 65536 Then
            
            If NewValue >= m_Min Then
            
                If NewValue <> m_Max Then
                
                    ' Set max value and update the range
                    m_Max = NewValue
                    
                    If m_hWnd Then
                        SendMessageW m_hWnd, PBM_SETRANGE, 0&, MakeDWord(m_Min, m_Max)
                    End If
                
                    ' Notify parent object
                    PropertyChanged "Max"
                    
                End If
            
            End If
            
        End If
        
    End If
    
End Property

Public Property Get Orientation() As OrientationConst
    ' Get current orientation
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal NewValue As OrientationConst)
    
    If NewValue <> m_Orientation Then
        
        ' Set the new value
        m_Orientation = NewValue
        
        ' Change the orientation
        CreateProgressBar
        
        ' Notify parent control
        PropertyChanged "Orientation"
        
    End If
    
End Property

Public Property Get Value() As Long
    ' Get value
    Value = m_Value
End Property

Public Property Let Value(ByVal NewValue As Long)
Attribute Value.VB_Description = "Returns or sets a control's current Value property."
Attribute Value.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
Attribute Value.VB_MemberFlags = "200"
    
    If NewValue >= m_Min Then
    
        If NewValue <= m_Max Then
    
            If NewValue <> m_Value Then
                
                ' Set value
                m_Value = NewValue
                
                If m_hWnd Then
                    SendMessageW m_hWnd, PBM_SETPOS, m_Value, 0&
                End If
                
                ' Notify parent object
                PropertyChanged "Value"
                
            End If
            
        End If
        
    End If
    
End Property
