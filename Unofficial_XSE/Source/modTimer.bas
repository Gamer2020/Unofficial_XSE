Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SetTimer Lib "user32" (ByVal Hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal Hwnd As Long, ByVal nIDEvent As Long) As Long

Private lTimerId As Long

' Stops the timer routine
Public Function EndTimer() As Boolean
    
    If lTimerId Then
        lTimerId = KillTimer(0&, lTimerId)
        lTimerId = 0
        EndTimer = True
    End If
    
End Function

' Starts the continuous calling of a private routine at a specific time interval.
Public Sub StartTimer(ByVal lInterval As Long)
    
    If lTimerId Then
        'End Current Timer
        EndTimer
    End If
    
    lTimerId = SetTimer(0&, 0&, ByVal lInterval, AddressOf TimerRoutine)
End Sub

' Routine which is called repeatedly by the timer API.
' Inputs are automatically generated.
Private Sub TimerRoutine(ByVal lHwnd As Long, ByVal lMsg As Long, ByVal lIDEvent As Long, ByVal lTime As Long)
    'Place your code here...
End Sub

