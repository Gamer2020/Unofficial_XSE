VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHexBox"
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

Private Const vbKeyCtrlC = vbKeyC - 64
Private Const vbKeyCtrlV = vbKeyV - 64
Private Const vbKeyCtrlX = vbKeyX - 64
Private Const vbKeyCtrlZ = vbKeyZ - 64

Private Const GWL_STYLE As Long = -16&
Private Const ES_UPPERCASE As Long = &H8&

Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public WithEvents TextBox As TextBox
Attribute TextBox.VB_VarHelpID = -1

Public Sub MakeUpperCase()
    SetWindowLongA TextBox.hWnd, GWL_STYLE, GetWindowLongA(TextBox.hWnd, GWL_STYLE) Or ES_UPPERCASE
End Sub

Private Function IsHex(ByRef HexString As String) As Boolean
Dim lTemp As Long
    
    On Error GoTo NotHex
    
    If LenB(HexString) Then
        lTemp = CLng("&H" & HexString)
        IsHex = True
    End If
    Exit Function
    
NotHex:
End Function

Private Sub TextBox_Change()
    
    ' If it's not a valid input
    If IsHex(TextBox.Text) = False Then
    
        ' Reset the textbox
        TextBox.Text = vbNullString
        Beep
        Exit Sub
        
    End If
    
End Sub

Private Sub TextBox_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        
        Case vbKeyBack, vbKeyCtrlC, vbKeyCtrlV, vbKeyCtrlX, vbKeyCtrlZ
            ' In these cases, nothing to do
    
        Case vbKey0 To vbKey9, vbKeyA To vbKeyF
            ' Numbers and A-F are OK
            
        Case vbKeyA + 32 To vbKeyF + 32
            ' a-f become A-F
            KeyAscii = KeyAscii - 32
            
        Case Else
            ' All other letters are unwanted
            KeyAscii = 0
            Beep
        
    End Select
    
End Sub
