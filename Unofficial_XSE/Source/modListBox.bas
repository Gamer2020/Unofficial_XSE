Attribute VB_Name = "modListBox"
Option Explicit

Private Const LB_ADDSTRING = &H180
Private Const LB_SETHORIZONTALEXTENT = &H194

Private Const DT_CALCRECT = &H400
Private Const SM_CXVSCROLL = 2

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Sub AddItem(frm As Form, lst As ListBox, sNewItem As String)
Dim rcText As RECT
Dim lNewWidth As Long
Dim lCurrWidth As Long
Dim lSysScrollWidth As Long
    
    ' If there are no items, reset the tag
    If lst.ListCount = 0 Then
        lst.Tag = vbNullString
    End If
    
    ' Get the current width used
    If LenB(lst.Tag) <> 0 Then
        lCurrWidth = CLng(lst.Tag)
    End If
   
    ' Get the width of the system scrollbar
    lSysScrollWidth = GetSystemMetrics(SM_CXVSCROLL)
   
    ' Use DrawText/DT_CALCRECT to determine item length
    DrawTextEx frm.hDC, sNewItem, -1&, rcText, DT_CALCRECT, ByVal 0&
    lNewWidth = rcText.Right + lSysScrollWidth
   
    ' If this is wider than the current setting,
    ' tweak the list and save the new horizontal
    ' extent to the tag property
    If lNewWidth > lCurrWidth Then
        SendMessage lst.hWnd, LB_SETHORIZONTALEXTENT, lNewWidth, ByVal 0&
        lst.Tag = lNewWidth
    End If
   
    ' Add the items to the control
    SendMessage lst.hWnd, LB_ADDSTRING, 0&, ByVal sNewItem

End Sub

Public Sub ListBoxToolTip(lst As ListBox, ByVal Y As Single, Optional ByRef sDefaultToolTip As String = vbNullString)
Dim lIndex As Long

    lIndex = Y \ (lst.Parent.TextHeight(vbNullString) * Screen.TwipsPerPixelY)
    
    ' Index evaluation
    lIndex = lIndex + lst.TopIndex
    
    If lIndex < lst.ListCount Then
        lst.ToolTipText = lst.List(lIndex)
    Else
        lst.ToolTipText = sDefaultToolTip
    End If

End Sub

