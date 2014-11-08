Attribute VB_Name = "modMain"
Option Explicit

' Copyright © 2009 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Private Const sMyName As String = "modMain"

Private Type tagInitCommonControlsEx
   dwSize As Long
   dwICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public m_hMod As Long

Public Sub Main()
Const sThis = "Main"
Const ICC_USEREX_CLASSES As Long = &H200&
Dim iccex As tagInitCommonControlsEx
   
   On Error GoTo LocalHandler
   
   ' Fille the iccex structure
   iccex.dwSize = Len(iccex)
   iccex.dwICC = ICC_USEREX_CLASSES
   
   ' Load library to prevent crash on closing
   m_hMod = LoadLibraryW(StrPtr("shell32.dll"))
   
   ' Load ComCtl32 classes
   InitCommonControlsEx iccex
   
   ' Show the main form
   frmMain.Show
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
