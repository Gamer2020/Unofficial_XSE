Attribute VB_Name = "modTimers"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Private m_colItems As Collection

Public Sub AddTimer(ByRef pobjTimer As APITimer, ByVal plngInterval As Long)
    
    If m_colItems Is Nothing Then
        Set m_colItems = New Collection
    End If
    
    pobjTimer.ID = SetTimer(0&, 0&, plngInterval, AddressOf Timer_CBK)
    m_colItems.Add ObjPtr(pobjTimer), pobjTimer.ID & "K"
    
End Sub

Public Sub RemoveTimer(ByRef pobjTimer As APITimer)
    
    On Error GoTo ErrHandler
    
    m_colItems.Remove pobjTimer.ID & "K"
    KillTimer 0&, pobjTimer.ID
    pobjTimer.ID = 0&
    
    If m_colItems.Count = 0 Then
        Set m_colItems = Nothing
    End If
    
    Exit Sub
    
ErrHandler:
End Sub

Public Sub Timer_CBK(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal SysTime As Long)
Dim lPointer As Long
Dim objTimer As APITimer

    On Error GoTo ErrHandler
    
    lPointer = m_colItems.Item(idEvent & "K")
    
    Set objTimer = PtrObj(lPointer)
    objTimer.RaiseTimerEvent
    Set objTimer = Nothing
    
    Exit Sub
    
ErrHandler:
End Sub

Private Function PtrObj(ByVal Pointer As Long) As Object
Dim objObject As Object

    RtlMoveMemory objObject, Pointer, 4&
    Set PtrObj = objObject
    RtlMoveMemory objObject, 0&, 4&
    
End Function
