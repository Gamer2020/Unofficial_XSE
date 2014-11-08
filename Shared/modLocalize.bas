Attribute VB_Name = "modLocalize"
Option Explicit

' Copyright © 2009 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Private Const sMyName = "modLocalize"
Private Const sNewFont As String = "Calibri"
Private Const sMenu As String = "Menu"
Private fCheckedAlready As Boolean
Private fNewFontInstalled As Boolean

Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Private Declare Function LoadStringW Lib "user32" (ByVal hInstance As Long, ByVal uId As Long, ByVal lpBuffer As Long, ByVal nBufferMax As Long) As Long

Public Sub Localize(ByRef frm As Form)
Const sThis = "Localize"
Dim ctl As Control
Dim sCtlType As String
Dim lVal As Long
Dim i As Long
    
    On Error GoTo LocalHandler
    
    ' Make sure the font installation wasn't checked already
    If fCheckedAlready = False Then
        
        ' Loop through all the installed fonts
        For i = 0 To Screen.FontCount - 1&
            
            ' Check if the current entry match the font
            If InStrB(Screen.Fonts(i), sNewFont) Then
                
                ' Font is installed, exit loop
                fNewFontInstalled = True
                Exit For
                
            End If
            
        Next i
        
        ' Set the CheckedAlready flag
        fCheckedAlready = True
        
    End If
    
    ' Ignore errors for unsupported controls
    On Error Resume Next
    
    ' Retrieve the form tag
    lVal = CLng(Val(frm.Tag))
    
    ' Check if it's valid
    If lVal Then
        ' Set the form's caption
        frm.Caption = LoadResString(lVal)
    End If
    
    ' If the font is installed
    If fNewFontInstalled Then
        ' Change the default one
        frm.FontName = sNewFont
    End If
    
    ' Loop through the different controls on the form
    For Each ctl In frm.Controls
        
        ' Get the current control type
        sCtlType = TypeName(ctl)
        
        ' Check if it's a menu
        If sCtlType <> sMenu Then
            
            ' Get the tag
            lVal = CLng(Val(ctl.Tag))
            
            ' If the font is installed
            If fNewFontInstalled Then
                
                If (Val(ctl.HelpContextID) = 0&) Then
                    ' Change the default one
                    ctl.FontName = sNewFont
                End If
                
            End If
            
        Else
            
            ' It's a menu, so no Tag property
            lVal = CLng(Val(ctl.HelpContextID))

        End If
        
        ' Make sure the tag is valid
        If lVal Then
            ' Set the localized caption
            ctl.Caption = LoadResString(lVal)
        End If
        
    Next
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

Public Function LoadString(ByVal id As Long) As String
Const sThis As String = "LoadString"
    
    On Error GoTo LocalHandler
    LoadString = LoadResString(id)
    Exit Function

LocalHandler:
    
    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
        
End Function
