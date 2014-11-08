VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCombomatic"
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

Private IsDeleting As Boolean
Private hWndEdit As Long
Private m_PrevIndex As Long

Private Declare Function FindWindowExA Lib "user32" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessageA Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_GETTEXT = &HD&
Private Const WM_GETTEXTLENGTH = &HE&
Private Const CB_GETEDITSEL = &H140&
Private Const CB_SETEDITSEL = &H142&
Private Const CB_GETCURSEL = &H147&
Private Const CB_SETCURSEL = &H14E&
Private Const CB_FINDSTRING = &H14C&
Private Const CB_FINDSTRINGEXACT = &H158&

Public WithEvents ComboBox As ComboBox
Attribute ComboBox.VB_VarHelpID = -1

Private Function LoWord(lDWord As Long) As Integer
    If lDWord And &H8000& Then
        LoWord = lDWord Or &HFFFF0000
    Else
        LoWord = lDWord And &HFFFF&
    End If
End Function

Private Function HiWord(lDWord As Long) As Integer
    HiWord = (lDWord And &HFFFF0000) \ &H10000
End Function

Private Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeDWord = (CLng(HiWord) * &H10000) Or (LoWord And &HFFFF&)
End Function

Private Sub GetEditHandle(hWnd As Long)
    hWndEdit = FindWindowExA(hWnd, 0&, "EDIT", vbNullString)
End Sub

Private Function GetTextLength() As Long
    GetTextLength = SendMessageA(hWndEdit, WM_GETTEXTLENGTH, 0&, ByVal 0&)
End Function

Private Function GetText() As String
Dim lLen As Long
    
    lLen = GetTextLength
    
    If lLen Then
        lLen = lLen + 1&
        GetText = Space$(lLen)
        SendMessageA hWndEdit, WM_GETTEXT, lLen, ByVal GetText
    End If

End Function

Private Function GetListIndex(hWnd As Long) As Long
    GetListIndex = SendMessageA(hWnd, CB_GETCURSEL, 0&, ByVal 0&)
End Function

Private Sub SetListIndex(hWnd As Long, lIndex As Long)
    SendMessageA hWnd, CB_SETCURSEL, lIndex, ByVal 0&
End Sub

Private Function GetSelStart(hWnd As Long) As Long
    GetSelStart = LoWord(SendMessageA(hWnd, CB_GETEDITSEL, 0&, ByVal 0&))
End Function

Private Function GetSelLength(hWnd As Long) As Long
Dim lWord As Long
Dim lLo As Long
Dim lHi As Long
    
    lWord = SendMessageA(hWnd, CB_GETEDITSEL, 0&, ByVal 0&)
    lLo = LoWord(lWord)
    lHi = HiWord(lWord)
    
    GetSelLength = lHi - lLo
    
End Function

Private Sub SetSel(hWnd As Long, lSelStart As Long, lSelLength As Long)
    SendMessageA hWnd, CB_SETEDITSEL, lSelStart, ByVal MakeDWord(lSelStart, lSelLength)
End Sub

Public Sub Change(hWnd As Long)
Dim lRet As Long
Dim lSelStart As Long
Dim lSelEnd As Long
    
    If IsDeleting Then
        IsDeleting = False
        Exit Sub
    End If
    
    If hWndEdit = 0& Then
        GetEditHandle hWnd
    End If
    
    lSelStart = GetSelStart(hWnd)
    lRet = SendMessageA(hWnd, CB_FINDSTRING, -1&, ByVal GetText)
    
    If lRet <> -1& Then
        
        SetListIndex hWnd, lRet
        SetSel hWnd, lSelStart, GetTextLength
        m_PrevIndex = lRet
        
    Else
        SetListIndex hWnd, m_PrevIndex
    End If
    
End Sub

Public Sub GotFocus(hWnd As Long)

    If hWndEdit = 0& Then
        GetEditHandle hWnd
    End If

    SetSel hWnd, 0, GetTextLength
    
End Sub

Public Sub LostFocus(hWnd As Long)
    Change hWnd
End Sub

Public Sub KeyDown(hWnd As Long, KeyCode As Integer, Shift As Integer)
Dim lRet As Long
    
    Select Case KeyCode
          
        Case vbKeyReturn
        
            KeyCode = 0
            
            If hWndEdit = 0& Then
                GetEditHandle hWnd
            End If
            
            If GetTextLength > GetSelLength(hWnd) Then
                lRet = SendMessageA(hWnd, CB_FINDSTRINGEXACT, -1&, ByVal Left$(GetText, GetTextLength - GetSelLength(hWnd)))
            Else
                lRet = -1
            End If
        
            If lRet <> -1 Then
                SetListIndex hWnd, lRet
            End If

            SetSel hWnd, GetTextLength, 0&
            
    End Select
    
End Sub

Private Sub ComboBox_Change()
    Change ComboBox.hWnd
End Sub

Private Sub ComboBox_GotFocus()
    GotFocus ComboBox.hWnd
End Sub

Private Sub ComboBox_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyDown ComboBox.hWnd, KeyCode, Shift
End Sub

Private Sub ComboBox_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyBack Then
        IsDeleting = True
    End If
    
End Sub

Private Sub ComboBox_LostFocus()
   LostFocus ComboBox.hWnd
End Sub
