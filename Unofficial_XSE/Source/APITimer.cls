VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "APITimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event Timer()

Private m_lTimerID As Long
Private m_lInterval As Long

Friend Property Let ID(ByVal lNewValue As Long)
     m_lTimerID = lNewValue
End Property

Friend Property Get ID() As Long
    ID = m_lTimerID
End Property

Public Sub StartTimer(ByVal Interval As Long)
    If m_lInterval <> Interval Then
        RemoveTimer Me
        AddTimer Me, Interval
        m_lInterval = Interval
    End If
End Sub

Public Sub StopTimer()
    If m_lTimerID > 0 Then
        RemoveTimer Me
    End If
End Sub

Private Sub Class_Terminate()
    StopTimer
End Sub

Friend Sub RaiseTimerEvent()
    RaiseEvent Timer
End Sub
