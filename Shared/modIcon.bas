Attribute VB_Name = "modIcon"
Option Explicit

' Copyright © 2009 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Private Const sMyName As String = "modIcon"

Private Const WM_SETICON As Long = &H80
Private Const SM_CXICON As Long = 11&
Private Const SM_CYICON As Long = 12&
Private Const SM_CXSMICON As Long = 49&
Private Const SM_CYSMICON As Long = 50&
Private Const LR_SHARED As Long = &H8000&
Private Const IMAGE_ICON As Long = 1&
Private Const ICON_SMALL As Long = 0&
Private Const ICON_BIG As Long = 1&
Private Const GW_OWNER As Long = 4&

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function LoadImageW Lib "user32" (ByVal hInst As Long, ByVal lpsz As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Sub SetIcon(ByVal hWnd As Long, ByVal IconResName As String, Optional ByVal SetAsAppIcon As Boolean = True)
Const sThis As String = "SetIcon"
Dim lhWndTop As Long
Dim lhWnd As Long
Dim cx As Long
Dim cy As Long
Dim hIconLarge As Long
Dim hIconSmall As Long
    
    On Error GoTo LocalHandler
    
    If SetAsAppIcon Then
        
        ' Find VB's hidden parent window:
        lhWnd = hWnd
        lhWndTop = lhWnd
      
        Do While lhWnd
         
            lhWnd = GetWindow(lhWnd, GW_OWNER)
         
            If lhWnd Then
                lhWndTop = lhWnd
            End If
            
        Loop
      
    End If
   
    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
    hIconLarge = LoadImageW(App.hInstance, StrPtr(IconResName), IMAGE_ICON, cx, cy, LR_SHARED)
   
    If SetAsAppIcon Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
    End If
   
    SendMessageLong hWnd, WM_SETICON, ICON_BIG, hIconLarge
   
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    hIconSmall = LoadImageW(App.hInstance, StrPtr(IconResName), IMAGE_ICON, cx, cy, LR_SHARED)
   
    If SetAsAppIcon Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
    End If
   
    SendMessageLong hWnd, WM_SETICON, ICON_SMALL, hIconSmall
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
   
End Sub
