Attribute VB_Name = "modTiming"
Option Explicit

Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private m_lT As Long

Public Sub StartTiming()
    timeBeginPeriod 1
    m_lT = timeGetTime
End Sub

Public Function EndTiming() As Long
    EndTiming = (timeGetTime - m_lT) + 1
    timeEndPeriod 1
End Function
