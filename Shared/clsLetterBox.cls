VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLetterBox"
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

Private Function IsLetter(ByRef Text As String) As Boolean
Dim bLetters() As Byte: bLetters = " ABCDEFGHIJKLMNOPRQSTUVWXYZabcdefghijklmnoprqstuvwxyz"
Dim bTemp() As Byte
Dim bTemp2(0) As Byte
Dim i As Long
    
    ' Ensure the text is not empty
    If LenB(Text) Then
        
        ' Convert the text into a byte array
        bTemp = Text
        
        ' Assume the text is valid
        IsLetter = True
        
        ' Loop through the letters
        For i = 0 To Len(Text) - 1&
            
            ' Get the current letter
            bTemp2(0) = bTemp(i * 2&)
            
            ' Check if it's valid
            If InStrB(bLetters, bTemp2) = 0& Then
            
                ' If not exit
                IsLetter = False
                Exit For
                
            End If
            
        Next i
        
    End If
    
End Function

Private Sub TextBox_Change()
    
    ' If it's not a valid input
    If IsLetter(TextBox.Text) = False Then
        
        ' Reset the textbox
        TextBox.Text = vbNullString
        Beep
        Exit Sub
        
    End If
    
End Sub

Private Sub TextBox_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
   
        Case vbKeyBack, vbKeyC - 64, vbKeyV - 64, vbKeyX - 64, vbKeyZ - 64
            ' Don't filter these keys
       
        Case vbKeyA + 32 To vbKeyZ + 32, vbKeyA To vbKeyZ, vbKeySpace
            ' Letters a-Z, A-Z and space are allowed
            
        Case Else
            ' Anything else is unwanted
            KeyAscii = 0
            Beep

    End Select
    
End Sub
