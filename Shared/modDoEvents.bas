Attribute VB_Name = "modDoEvents"
Option Explicit

' Copyright © 2009 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type MSG
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Const PM_REMOVE As Long = &H1&

Private Declare Function PeekMessageW Lib "user32" (ByRef lpMsg As MSG, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (ByRef lpMsg As MSG) As Long
Private Declare Function DispatchMessageW Lib "user32" (ByRef lpMsg As MSG) As Long

' Original written by Nir Sofer
' http://www.nirsoft.net
Public Sub MyDoEvents()
Dim CurrMsg As MSG
    
    ' The following loop extract all messages from the queue
    ' and dispatch them to the appropriate window
    Do While PeekMessageW(CurrMsg, 0&, 0&, 0&, PM_REMOVE)
        TranslateMessage CurrMsg
        DispatchMessageW CurrMsg
    Loop
    
End Sub
