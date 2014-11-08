Attribute VB_Name = "modMsgBox"
Option Explicit

' Copyright © 2009 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Private Const sMyName = "modMsgBox"

' The operating system is Windows Server 2008, Windows Vista,
' Windows Server 2003, Windows XP, or Windows 2000.
Private Const VER_PLATFORM_WIN32_NT = 2&
Private Const VER_MAJOR_2K = 5&
Private Const VER_MAJOR_VISTA = 6&

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128&
End Type

' API Declares
' ---
Private Declare Function GetVersionExA Lib "kernel32" (ByRef lpVersionInfo As OSVERSIONINFO) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function MessageBoxW Lib "user32" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
Private Declare Function TaskDialog Lib "comctl32" (ByVal hwndParent As Long, ByVal hInstance As Long, ByVal pszWindowTitle As Long, ByVal pszMainInstruction As Long, ByVal pszContent As Long, ByVal dwCommonButtons As Long, ByVal pszIcon As Long, pnButton As Long) As Long

' MessageBox Button constants
' ---
Private Const MB_OK = &H0&
' The message box contains one push button: OK. This is the default.
Private Const MB_OKCANCEL = &H1&
' The message box contains two push buttons: OK and Cancel.
Private Const MB_ABORTRETRYIGNORE = &H2&
' The message box contains three push buttons: Abort, Retry, and Ignore.
Private Const MB_YESNOCANCEL = &H3&
' The message box contains three push buttons: Yes, No, and Cancel.
Private Const MB_YESNO = &H4&
' The message box contains two push buttons: Yes and No.
Private Const MB_RETRYCANCEL = &H5&
' The message box contains two push buttons: Retry and Cancel.
Private Const MB_CANCELTRYCONTINUE = &H6&
' Microsoft Windows 2000/XP: The message box contains three push buttons:
' Cancel, Try Again, Continue. Use this message box type instead of MB_ABORTRETRYIGNORE.
Private Const MB_HELP = &H4000&
' Windows 95/98/Me, Windows NT 4.0 and later: Adds a Help button to the message box.
' When the user clicks the Help button or presses F1, the system sends a WM_HELP
' message to the owner.


' MessageBox Icon constants
' ---
Private Const MB_NO_ICON = &H0&
' No icon appears in the message box. This is the default.
Private Const MB_ICONERROR = &H10&
Private Const MB_ICONHAND = MB_ICONERROR
Private Const MB_ICONSTOP = MB_ICONERROR
' A stop-sign icon appears in the message box.
Private Const MB_ICONQUESTION = &H20&
' A question-mark icon appears in the message box.
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_ICONWARNING = MB_ICONEXCLAMATION
' An exclamation-point icon appears in the message box.
Private Const MB_ICONINFORMATION = &H40&
Private Const MB_ICONASTERISK = MB_ICONINFORMATION
' An icon consisting of a lowercase letter i in a circle appears in the message box.


' MessageBox Default Button constants
' ---
Private Const MB_DEFBUTTON1 = &H0&
' The first button is the default button.
' MB_DEFBUTTON1 is the default unless MB_DEFBUTTON2, MB_DEFBUTTON3, or
' MB_DEFBUTTON4 is specified.
Private Const MB_DEFBUTTON2 = &H100&
' The second button is the default button.
Private Const MB_DEFBUTTON3 = &H200&
' The third button is the default button.
Private Const MB_DEFBUTTON4 = &H300&
' The fourth button is the default button.


' MessageBox Modal constants
' ---
Private Const MB_APPLMODAL = &H0&
' The user must respond to the message box before continuing work in the window
' identified by the hWnd parameter. However, the user can move to the windows of other
' threads and work in those windows. Depending on the hierarchy of windows in the
' application, the user may be able to move to other windows within the thread.
' All child windows of the parent of the message box are automatically disabled,
' but pop-up windows are not. MB_APPLMODAL is the default if neither MB_SYSTEMMODAL
' nor MB_TASKMODAL is specified.
Private Const MB_SYSTEMMODAL = &H1000&
' Same as MB_APPLMODAL except that the message box has the WS_EX_TOPMOST style.
' Use system-modal message boxes to notify the user of serious, potentially damaging
' errors that require immediate attention (for example, running out of memory).
' This flag has no effect on the user's ability to interact with windows other than
' those associated with hWnd.
Private Const MB_TASKMODAL = &H2000&
' Same as MB_APPLMODAL except that all the top-level windows belonging to the current
' thread are disabled if the hWnd parameter is NULL. Use this flag when the calling
' application or library does not have a window handle available but still needs
' to prevent input to other windows in the calling thread without suspending other threads.


' MessageBox Special constants
' ---
Private Const MB_SETFOREGROUND = &H10000
' The message box becomes the foreground window. Internally, the system calls the
' SetForegroundWindow function for the message box.
Private Const MB_DEFAULT_DESKTOP_ONLY = &H20000
' Windows NT 4.0 and earlier: If the current input desktop is not the default desktop,
' MessageBox fails.
' Windows 2000/XP: If the current input desktop is not the default desktop, MessageBox does not
' return until the user switches to the default desktop.
' Windows 95/98/Me: This flag has no effect.
Private Const MB_SERVICE_NOTIFICATION_NT3X = &H40000
' Windows NT/2000/XP: This value corresponds to the value defined for MB_SERVICE_NOTIFICATION
' for Windows NT version 3.51.
Private Const MB_TOPMOST = &H40000
' The message box is created with the WS_EX_TOPMOST window style.
Private Const MB_RIGHT = &H80000
' The text is right-justified.
Private Const MB_RTLREADING = &H100000
' Displays message and caption text using right-to-left reading order on Hebrew and Arabic systems.
Private Const MB_SERVICE_NOTIFICATION = &H200000
' Windows NT/2000/XP: The caller is a service notifying the user of an event. The function displays
' a message box on the current active desktop, even if there is no user logged on to the computer.
' Terminal Services: If the calling thread has an impersonation token, the function directs the message
' box to the session specified in the impersonation token.
' If this flag is set, the hWnd parameter must be NULL. This is so that the message box can appear on
' a desktop other than the desktop corresponding to the hWnd.


' TaskDialog Button constants
' ---
Private Const TDCBF_OK_BUTTON = &H1&
' The task dialog contains the push button: OK.
Private Const TDCBF_YES_BUTTON = &H2&
' The task dialog contains the push button: Yes.
Private Const TDCBF_NO_BUTTON = &H4&
' The task dialog contains the push button: No.
Private Const TDCBF_CANCEL_BUTTON = &H8&
' The task dialog contains the push button: Cancel.
' This button must be specified for the dialog box to respond to typical cancel
' actions (Alt-F4 and Escape).
Private Const TDCBF_RETRY_BUTTON = &H10&
' The task dialog contains the push button: Retry.
Private Const TDCBF_CLOSE_BUTTON = &H20&
' The task dialog contains the push button: Close.


' TaskDialog Icon constants
' ---
Private Const TD_NO_ICON = 0&
' No icon appears in the task dialog. This is the default.
Private Const TD_WARNING_ICON = 65535
' An exclamation-point icon appears in the task dialog.
Private Const TD_ERROR_ICON = 65534
' A stop-sign icon appears in the task dialog.
Private Const TD_INFORMATION_ICON = 65533
' An icon consisting of a lowercase letter i in a circle appears in the task dialog.
Private Const TD_SHIELD_ICON = 65532
' A shield icon appears in the task dialog.
Private Const TD_SHIELD_BLUE_ICON = 65531
' A shield icon on a blue background appears in the task dialog.
Private Const TD_SHIELD_WARNING_ICON = 65530
' An icon consisting of an exclamation-point in a shield appears in the task dialog.
Private Const TD_SHIELD_ERROR_ICON = 65529
' An icon consisting of a stop-sign in a shield appears in the task dialog.
Private Const TD_SHIELD_SUCCESS_ICON = 65528
' An icon consisting of a tick-sign in a shield appears in the task dialog.
Private Const TD_SHIELD_BROWN_ICON = 65527
' A shield icon on a brown background appears in the task dialog.


' Return Value constants
' ---
Private Const IDFAIL = 0&
' The function failed.
Private Const IDOK = 1&
' The OK button was selected.
Private Const IDCANCEL = 2&
' The Cancel button was selected.
Private Const IDABORT = 3&
' The Abort button was selected.
Private Const IDRETRY = 4&
' The Retry button was selected.
Private Const IDIGNORE = 5&
' The Ignore button was selected.
Private Const IDYES = 6&
' The Yes button was selected.
Private Const IDNO = 7&
' The No button was selected.
Private Const IDCLOSE = 8&
' The Close button was selected.
Private Const IDHELP = 9&
' The Help button was selected.
Private Const IDTRYAGAIN = 10&
' The Try Again button was selected.
Private Const IDCONTINUE = 11&
' The Continue button was selected.


' Enums
' ---
Public Enum MessageButtons
    vbOK = MB_OK
    vbOKCancel = MB_OKCANCEL
    vbAbortRetryIgnore = MB_ABORTRETRYIGNORE
    vbYesNoCancel = MB_YESNOCANCEL
    vbYesNo = MB_YESNO
    vbRetryCancel = MB_RETRYCANCEL
    vbCancelTryContinue = MB_CANCELTRYCONTINUE
    vbCloseButton = TDCBF_CLOSE_BUTTON
    vbHelpButton = MB_HELP
End Enum

Public Enum MessageIcon
    vbNone = MB_NO_ICON
    vbError = MB_ICONERROR
    vbQuestion = MB_ICONQUESTION
    vbWarning = MB_ICONEXCLAMATION
    vbInformation = MB_ICONINFORMATION
    vbShield = TD_SHIELD_ICON
    vbShieldBlue = TD_SHIELD_BLUE_ICON
    vbShieldWarning = TD_SHIELD_WARNING_ICON
    vbShieldError = TD_SHIELD_ERROR_ICON
    vbShieldSuccess = TD_SHIELD_SUCCESS_ICON
    vbShieldBrown = TD_SHIELD_BROWN_ICON
End Enum

Public Enum MessageOptions
    vbDefaultButton1 = MB_DEFBUTTON1
    vbDefaultButton2 = MB_DEFBUTTON2
    vbDefaultButton3 = MB_DEFBUTTON3
    vbDefaultButton4 = MB_DEFBUTTON4
    vbAppModal = MB_APPLMODAL
    vbSystemModal = MB_SYSTEMMODAL
    vbTaskModal = MB_TASKMODAL
    vbSetForeground = MB_SETFOREGROUND
    vbDefaultDesktopOnly = MB_DEFAULT_DESKTOP_ONLY
    vbServiceNotificationNT3x = MB_SERVICE_NOTIFICATION_NT3X
    vbTopMost = MB_TOPMOST
    vbRightAlign = MB_RIGHT
    vbRtlReading = MB_RTLREADING
    vbServiceNotification = MB_SERVICE_NOTIFICATION
End Enum

Public Enum MessageResult
    vbFail = IDFAIL
    vbOK = IDOK
    vbCancel = IDCANCEL
    vbAbort = IDABORT
    vbRetry = IDRETRY
    vbIgnore = IDIGNORE
    vbYes = IDYES
    vbNo = IDNO
    vbClose = IDCLOSE
    vbHelp = IDHELP
    vbTryAgain = IDTRYAGAIN
    vbContinue = IDCONTINUE
End Enum

Private lOSVersion As Long

Private Function GetVersion() As Long
Dim osv As OSVERSIONINFO
    
    If lOSVersion = 0& Then
    
        ' Set the size
        osv.dwOSVersionInfoSize = Len(osv)
        
        ' If the function succeeds
        If GetVersionExA(osv) Then
            ' Check if the platform is NT
            If osv.dwPlatformId = VER_PLATFORM_WIN32_NT Then
                ' Return the Major version
                lOSVersion = osv.dwMajorVersion
            End If
        End If
    
    End If
    
    GetVersion = lOSVersion

End Function

' Override the default MsgBox function
Public Function MsgBox(ByVal sPrompt As String, Optional ByVal lButtons As Long, Optional ByVal sTitle As String = vbNullString) As MessageResult
    MsgBox = MessageBox(sPrompt, sTitle, lButtons And &HF&, lButtons And &HF0&, lButtons And &HFFFF00)
End Function

Public Function MessageBox(ByVal sPrompt As String, Optional ByVal sTitle As String = vbNullString, Optional ByVal lButtons As MessageButtons = vbOKOnly, Optional ByVal lIcon As MessageIcon = vbNone, Optional ByVal lOptions As MessageOptions, Optional ByVal fForceLegacy As Boolean = False) As MessageResult
Const sThis = "MessageBox"
Dim sMainPrompt As String
Dim sContent As String
Dim dwButtons As Long
Dim lStyle As Long
Dim lPos As Long
Dim lRet As Long
    
    On Error GoTo LocalHandler
    
    ' If not title was specified
    If StrPtr(sTitle) = 0 Then
        ' Use the app one
        sTitle = App.Title
    End If
    
    ' If the OS is not Vista or higher or the ForceLegacy flag was enabled
    If GetVersion < VER_MAJOR_VISTA Or fForceLegacy = True Then
        
        ' Check if the Help button is enabled
        If (lButtons And vbHelpButton) = vbHelpButton Then
            ' If so, set the style accordingly
            lStyle = vbHelpButton
        End If
        
        ' Mask the Buttons and the Icon
        ' This is needed to prevent possible high values
        lButtons = lButtons And &HF&
        lIcon = lIcon And &HFFFF&
        
        ' The Shields are not available for MessageBoxes
        ' Therefore we need to adjust the icon
        Select Case lIcon
            Case vbShield, vbShieldBlue, vbShieldSuccess, vbShieldBrown
                lIcon = vbInformation
            Case vbShieldWarning
                lIcon = vbWarning
            Case vbShieldError
                lIcon = vbError
        End Select
        
        ' Mask the Options
        lOptions = lOptions And &HFFFF00
        
        ' If the Buttons aren't vbCancelTryContinue
        If (lButtons And vbCancelTryContinue) <> vbCancelTryContinue Then
            ' Set the style normally
            lStyle = lStyle Or lButtons Or lIcon Or lOptions
        Else
            ' Check if the OS is 2000+
            If GetVersion >= VER_MAJOR_2K Then
                lStyle = lStyle Or (lButtons Or lIcon Or lOptions)
            Else
                ' vbCancelTryContinue is not supported
                ' Replace it with vbAbortRetryIgnore
                lStyle = lStyle Or (vbAbortRetryIgnore Or lIcon Or lOptions)
            End If
        End If
    
        ' Display the MessageBox
        lRet = MessageBoxW(GetActiveWindow, StrPtr(sPrompt), StrPtr(sTitle), lStyle)
        
        ' If the buttons were vbCancelTryContinue
        If (lButtons And vbCancelTryContinue) = vbCancelTryContinue Then
            ' If the OS isn't 2000+
            If GetVersion < VER_MAJOR_2K Then
                ' Adjust the return value
                Select Case lRet
                    Case vbRetry, vbIgnore
                        ' Retry and Ignore become Try Again and Continue, respectively
                        lRet = lRet + 6&
                    Case vbAbort
                        ' Abort becomes Cancel
                        lRet = vbCancel
                End Select
            End If
        End If
        
    Else
        
        ' Since we're going to use a TaskDialog
        ' we need to adjust the values
        Select Case lButtons And &HF&
            Case vbOKOnly
                dwButtons = TDCBF_OK_BUTTON
            Case vbOKCancel
                dwButtons = TDCBF_OK_BUTTON Or TDCBF_CANCEL_BUTTON
            Case vbAbortRetryIgnore, vbCancelTryContinue
                dwButtons = TDCBF_CLOSE_BUTTON Or TDCBF_RETRY_BUTTON Or TDCBF_CANCEL_BUTTON
            Case vbYesNoCancel
                dwButtons = TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON Or TDCBF_CANCEL_BUTTON
            Case vbYesNo
                dwButtons = TDCBF_YES_BUTTON Or TDCBF_NO_BUTTON
        End Select
        
        ' Check if the Close button is enabled
        If (lButtons And vbCloseButton) = vbCloseButton Then
            ' If so, set the Buttons accordingly
            dwButtons = dwButtons Or vbCloseButton
        End If
        
        ' Mask the Icon
        lIcon = lIcon And &HFFFF&
        
        ' If the icon is not a shield one, adjust the value
        Select Case lIcon
            Case vbError
                lIcon = TD_ERROR_ICON
            Case vbWarning
                lIcon = TD_WARNING_ICON
            Case vbQuestion, vbInformation
                lIcon = TD_INFORMATION_ICON
        End Select
        
        ' If there's at least one New Line
        If InStrB(sPrompt, vbNewLine) <> 0 Then
            
            ' Get the position
            lPos = InStr(sPrompt, vbNewLine)
            ' Get the content part starting after vbNewLine
            sContent = Mid$(sPrompt, lPos + 2&)
        
        'Or a Carriage Return
        ElseIf InStrB(sPrompt, vbCr) <> 0 Then
        
            ' Get the position
            lPos = InStr(sPrompt, vbCr)
            ' Get the content part starting after vbCr
            sContent = Mid$(sPrompt, lPos + 1&)
            
        ' Or a Line Feed
        ElseIf InStrB(sPrompt, vbLf) <> 0 Then
            
            ' Get the position
            lPos = InStr(sPrompt, vbLf)
            ' Get the content part starting after vbLf
            sContent = Mid$(sPrompt, lPos + 1&)
            
        End If
        
        ' If the prompt was splitted
        If lPos <> 0 Then
            ' Strip the content part from the prompt
            sPrompt = Left$(sPrompt, lPos - 1&)
        End If
    
        ' Display the TaskDialog
        TaskDialog GetActiveWindow, 0&, StrPtr(sTitle), StrPtr(sPrompt), StrPtr(sContent), dwButtons, lIcon, lRet
        
        ' See if we need to adjust the return value
        Select Case lButtons And &HF&
            Case vbAbortRetryIgnore
                Select Case lRet
                    Case vbCancel
                        ' Cancel becomes Ignore
                        lRet = vbIgnore
                    Case vbClose
                        ' Close becomes Abort
                        lRet = vbAbort
                End Select
            Case vbCancelTryContinue
                Select Case lRet
                    Case vbRetry
                        ' Retry becomes Try Again
                        lRet = vbTryAgain
                    Case vbClose
                        ' Close becomes Cancel
                        lRet = vbCancel
                    Case vbCancel
                        ' Cancel becomes Continue
                        lRet = vbContinue
                End Select
        End Select
        
    End If
    
    MessageBox = lRet
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
