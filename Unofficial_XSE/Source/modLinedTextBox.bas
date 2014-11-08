Attribute VB_Name = "modLinedTextBox"
Option Explicit

' *********************************************************************************
'  modLinedTextBox - a module to add line numbers to the common textbox.
' *********************************************************************************
'
'   Author: G. D. Sever   (garrett@elitevb.com)
'     Date: Jan 2004
'     Mods:
'
'     Desc: Module that allows a user to add and remove line numbers from the
'            non-client area of a TextBox or RichTextBox control. The line numbers
'            are printed in the same font and size as the font in the textbox, or
'            or in a standardized font for the RichTextBox.
'
'           This module should adjust for all combinations of border styles and
'            appearances (flat or none / fixed single / 3D). It also makes some
'            subtle adjustments depending on whether its a RTB or a standard
'            textbox since they behave slightly different.
'

' Windows Messages
Private Enum WindowsMessages
    WM_DESTROY = &H2
    WM_SETFOCUS = &H7
    WM_KILLFOCUS = &H8
    WM_PAINT = &HF
    'WM_ERASEBKGND = &H14
    WM_GETFONT = &H31
    WM_NCCALCSIZE = &H83
    WM_NCPAINT = &H85
    WM_NCHITTEST = &H84
    WM_KEYDOWN = &H100
    WM_KEYUP = &H101
    WM_CHAR = &H102
    WM_CTLCOLOREDIT = &H133
    WM_MOUSEMOVE = &H200
    WM_LBUTTONDOWN = &H201
    WM_LBUTTONUP = &H202
    WM_PASTE = &H302
    WM_USER = &H400
End Enum

' Constant used with WM_MOUSEMOVE to determine button states.
Private Const MK_LBUTTON = &H1

' Edit messages
Private Enum EditMessages
    EM_GETFIRSTVISIBLELINE = &HCE
    EM_GETLINECOUNT = &HBA
    EM_GETSEL = &HB0
    EM_SETSEL = &HB1
    EM_LINEINDEX = &HBB                            ' Gets index of the first character in a specified line #
    EM_LINEFROMCHAR = &HC9
    EM_POSFROMCHAR = (WM_USER + 38)
    EM_CHARFROMPOS = &HD7
End Enum

' Draw text items
Private Enum DrawTextConstatns
    DT_CENTER = &H1
    DT_right = &H2
    DT_VCENTER = &H4
    DT_SINGLELINE = &H20
    DT_CALCRECT = &H400
End Enum

Private Enum SetWindowPosConstants
    SWP_ASYNCWINDOWPOS = &H4000
    SWP_DEFERERASE = &H2000
    SWP_FRAMECHANGED = &H20
    SWP_HIDEWINDOW = &H80
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200
    SWP_NOREDRAW = &H8
    SWP_NOREPOSITION = SWP_NOOWNERZORDER
    SWP_NOSENDCHANGING = &H400
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_SHOWWINDOW = &H40
End Enum

' ******************************************************************************************************
' ******************************************************************************************************
'
'  API User Defined Type (UDT) Declarations
'
' ******************************************************************************************************
' ******************************************************************************************************
' a point.
Private Type POINTAPI
    x                           As Long
    Y                           As Long
End Type

' A rectangle sucture
Private Type RECT
    Left                        As Long
    Top                         As Long
    Right                       As Long
    Bottom                      As Long
End Type

' Window position UDT
Private Type WINDOWPOS
   hWnd                         As Long
   hWndInsertAfter              As Long
   x                            As Long
   Y                            As Long
   cx                           As Long
   cy                           As Long
   flags                        As Long
End Type

' Non-client calculate size parameters
Private Type NCCALCSIZE_PARAMS
   rgrc(0 To 2)                 As RECT ' rectangles defining window positions for the WM_NCCALCSIZE message
                                        '  0 = new window coordinates
                                        '  1 = coordinates of the window before it was moved or resized.
                                        '  2 = coordinates of the window's client area before the window was moved or resized.
   lppos                        As Long ' pointer to a WINDOWPOS UDT - size and position values specified in the operation that moved or resized the window
End Type

' ***************************************************************************************************************
' ***************************************************************************************************************
'
'  API Function and Sub Declarations
'
' ***************************************************************************************************************
' ***************************************************************************************************************

' APIs to install our subclassing routines
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' APIs used to keep track of process addresses
Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long

' Graphics declarations & misc
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal w As Long, ByVal e As Long, ByVal O As Long, ByVal w As Long, ByVal i As Long, ByVal u As Long, ByVal S As Long, ByVal c As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32.dll" () As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
'Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
'Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_WNDPROC = -4

Public lNumDigits As Long

Public Sub ShowLines(aTxt As TextBox, ByVal aBool As Boolean)
    ' Turn line numbers on and off.
    
    If aBool Then
        If GetProp(aTxt.hWnd, "OrigWndProc") = 0 Then
            ' Subclass the control to start
            SubclassLinedEdit aTxt
        End If
        lNumDigits = Len(CStr(SendMessage(aTxt.hWnd, EM_GETLINECOUNT, 0&, ByVal 0&)))
        If lNumDigits < 4 Then lNumDigits = 4
        SetProp aTxt.hWnd, "NumDigits", lNumDigits
    ElseIf Not aBool And GetProp(aTxt.hWnd, "OrigWndProc") <> 0 Then
        ' Unsubclass the control to stop the line number processes
        SetWindowLong aTxt.hWnd, GWL_WNDPROC, GetProp(aTxt.hWnd, "OrigWndProc")
        RemoveProp aTxt.hWnd, "OrigWndProc"
        RemoveProp aTxt.hWnd, "ControlPtr"
    End If
    ' Make sure the control is updated w/ correct non-client areas
    '  (where we draw the line numbers) by using this little SetWindowPos hack.
    SetWindowPos aTxt.hWnd, _
                 0&, 0&, 0&, 0&, 0&, _
                 SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOZORDER Or _
                 SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_FRAMECHANGED
    
    
End Sub

' ***************************************************************************************************
'
'  Private Functions and Subs- Used for custom behavior of our lined textbox.
'
' ***************************************************************************************************
Private Sub SubclassLinedEdit(aControl As TextBox)
    
    Dim origProc As Long
    
    ' Make sure there are no typos before subclassing.
    LinedEditProc 0, 0, 0, 0
   ' NotifyProc 0, 0, 0, 0, 0
    LineNumberProc 0, 0, 0, 0, 0
    PaintLineNumbers 0
    
    ' Make sure we're not already subclassing
    If GetProp(aControl.hWnd, "OrigWndProc") <> 0 Then Exit Sub
    
    ' Start subclassing
    origProc = SetWindowLong(aControl.hWnd, GWL_WNDPROC, AddressOf LinedEditProc)
    ' Store the process address for later
    SetProp aControl.hWnd, "OrigWndProc", origProc
    ' Save the Textbox's  pointer address
    SetProp aControl.hWnd, "ControlPtr", ObjPtr(aControl)

End Sub

Private Function LinedEditProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim origProc        As Long          ' address of the original window process for the control
        
    If hWnd = 0 Then Exit Function
    
    ' Get the original process address
    origProc = GetProp(hWnd, "OrigWndProc")

    
    If origProc <> 0 Then
        If uMsg = WM_DESTROY Then
            ' Either the control is being desoyed or something weird
            '  happened and we lost the pointer to the textbox. Unhook
            '  and invoke the original window procedure.
            SetWindowLong hWnd, GWL_WNDPROC, origProc
            ' Clean up our stored values
            RemoveProp hWnd, "OrigWndProc"
            RemoveProp hWnd, "ControlPtr"
            ' Invoke the original window procedure
            LinedEditProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        ElseIf uMsg = WM_NCCALCSIZE Then
            ' We need to resize the non-client area to accomodate the
            '  line numbers off to the left.
            LinedEditProc = LineNumberProc(origProc, hWnd, uMsg, wParam, lParam)
        'ElseIf uMsg = WM_ERASEBKGND Then
        '    LinedEditProc = 1
        ElseIf uMsg = WM_CTLCOLOREDIT Then
            LinedEditProc = 1
        ElseIf uMsg = WM_PAINT Or uMsg = WM_PASTE Then 'uMsg = WM_NCPAINT Or
            ' We'll make sure that the line numbers are redrawn when the
            '  control needs to redraw itself.
            LinedEditProc = LineNumberProc(origProc, hWnd, uMsg, wParam, lParam)
        ElseIf uMsg = WM_SETFOCUS Or uMsg = WM_KILLFOCUS Then
            ' Catching these messages allow the line numbers to be redrawn with the
            '  active/inactive color for selected rows.
            LinedEditProc = LineNumberProc(origProc, hWnd, uMsg, wParam, lParam)
        ElseIf uMsg = WM_KEYDOWN Or uMsg = WM_KEYUP Or uMsg = WM_CHAR Then
            ' The
            LinedEditProc = LineNumberProc(origProc, hWnd, uMsg, wParam, lParam)
        ElseIf uMsg = WM_LBUTTONDOWN Or uMsg = WM_LBUTTONUP Or _
               (uMsg = WM_MOUSEMOVE And wParam = MK_LBUTTON) Then
            ' We want the line numbers to redraw to reflect new selected areas.
            LinedEditProc = LineNumberProc(origProc, hWnd, uMsg, wParam, lParam)
        ElseIf uMsg = WM_NCHITTEST Then
            ' This allows us to respond to the user clicking in the line number
            '  "tray". By returning 1 instead of 0, we tell the caller that
            '  if the user clicked in "nowhere land", then it was actually
            '  the client area. This somewhat magically tricks the tray into
            '  behaving just like the client area in terms of selection and clicking.
            '  Pretty neat.
            LinedEditProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
            If LinedEditProc = 0 Then LinedEditProc = 1
        ElseIf uMsg = EM_SETSEL Then
            ' Occurs when "Select All" is chosen from the context menu.
            LinedEditProc = LineNumberProc(origProc, hWnd, uMsg, wParam, lParam)
        Else
            ' Invoke the original window procedure
            LinedEditProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        End If
    Else
        ' This is just incase something freaky happens and we lose the
        '  address of the old window procedure.
        LinedEditProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If

End Function

Private Function LineNumberProc(ByVal origProc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim aNCCS           As NCCALCSIZE_PARAMS ' UDT that allows us to resize the non-client area.
    Dim aWinPos         As WINDOWPOS         ' Window position sucture
    Dim aFont           As Long              ' handle to a font used for determining tray width
    Dim aDC             As Long              ' device context of the control
    Dim aPt             As POINTAPI          ' used to determine width of line number tray
    Dim lTrayWidth      As Long              ' width of the line number tray in the non-client area.
    Dim lNumDigits      As Long
    
    If hWnd = 0 Then Exit Function
    
    If uMsg = WM_NCCALCSIZE Then
        If wParam <> 0 Then
            ' Determine the width of the non-client area for the numbers from the font size...
            lNumDigits = GetProp(hWnd, "NumDigits")
            If lNumDigits < 4 Then lNumDigits = 4
            aDC = GetWindowDC(hWnd)
            aFont = SelectObject(aDC, SendMessage(hWnd, WM_GETFONT, ByVal 0&, ByVal 0&))
            GetTextExtentPoint32 aDC, String$(lNumDigits, "0"), lNumDigits, aPt
            lTrayWidth = aPt.x + 10 * 2 ' 10 pixel padding on both sides
            SetProp hWnd, "TrayWidth", lTrayWidth
            SelectObject aDC, aFont
            ReleaseDC hWnd, aDC
            ' Get the non-client calc size rectangle and window position
            RtlMoveMemory aNCCS, ByVal lParam, Len(aNCCS)
            RtlMoveMemory aWinPos, ByVal aNCCS.lppos, Len(aWinPos)
            ' Populate the non-client rectangle UDT information
            With aNCCS.rgrc(0)
                .Left = aWinPos.x
                .Top = aWinPos.Y
                .Right = aWinPos.x + aWinPos.cx
                .Bottom = aWinPos.Y + aWinPos.cy
            End With
            ' Make an adjustment for our line numbers
            aNCCS.rgrc(0).Left = aNCCS.rgrc(0).Left + lTrayWidth
            ' Duplicate these values in the other rectangle
            LSet aNCCS.rgrc(1) = aNCCS.rgrc(0)
            ' copy it back to the lParam pointer address so the process uses the information.
            RtlMoveMemory ByVal lParam, aNCCS, Len(aNCCS)
        End If
        ' invoke the default process
        LineNumberProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
    ElseIf (uMsg = WM_PAINT Or uMsg = WM_SETFOCUS Or uMsg = WM_KILLFOCUS) Then
        PaintLineNumbers hWnd
        ' invoke the default process
        LineNumberProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
    ElseIf uMsg = WM_NCPAINT Then
        PaintLineNumbers hWnd
        ' invoke the default process
        LineNumberProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
    ElseIf uMsg = WM_KEYDOWN Or uMsg = WM_KEYUP Or uMsg = WM_CHAR Then
        ' invoke the default process
        LineNumberProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        ' Paint the line numbers
        PaintLineNumbers hWnd
    ElseIf uMsg = WM_LBUTTONDOWN Or uMsg = WM_LBUTTONUP Or (uMsg = WM_MOUSEMOVE And wParam = MK_LBUTTON) Then
        ' invoke the default process
        LineNumberProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        ' Paint the line numbers
        PaintLineNumbers hWnd
    ElseIf uMsg = EM_SETSEL Or uMsg = WM_PASTE Then
        ' invoke the default process
        LineNumberProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
        ' Paint the line numbers
        PaintLineNumbers hWnd
    Else
        ' invoke the default process
        LineNumberProc = CallWindowProc(origProc, hWnd, uMsg, wParam, lParam)
    End If
    
End Function

Private Sub PaintLineNumbers(ByVal hWnd As Long)

    ' **************************************************************************
    '  Draws the line numbers off to the left of the textbox.
    ' **************************************************************************
    
    Dim backBuffDC      As Long     ' Back buffer device context
    Dim backBuffBmp     As Long     ' Back buffer bitmap
    Dim aBrush          As Long     ' Brush used to fill in the BG color
    Dim aRECT           As RECT     ' A rectangle.
    Dim aPt             As POINTAPI ' A point. Used to determine character indices
'    Dim lCoords         As Long     ' A long value representing x/y coordinates
    Dim curSel          As Long     ' Current selection's start index
    Dim curSelEnd       As Long     ' Current selection's ending index
    Dim curSelLine      As Long     ' Line # where current selection begins
    Dim curSelEndLine   As Long     ' Line # where current selection ends
    Dim lineTop         As Long     ' first line # displayed in the edit control
    Dim lineBottom      As Long     ' last line # displayed in the edit control
    Dim lineChar        As Long     ' 0 based index of first character in a specific line
    Dim TotalLines      As Long     ' total number of lines of text in the edit control
    Dim curLine         As Long     ' current line - used for looping thru
    Dim aPen            As Long     ' A pen to outline the line number
    Dim selRECT         As RECT     ' A RECT UDT for the selected lines box.
    Dim rcWindow        As RECT     ' window rectangle
    Dim aFont           As Long     ' The font we use to draw on our backbuffer
    Dim FontHeight      As Long     ' Height of the courier new font in pixels
    Dim HasFocus        As Boolean  ' Whether the control has focus or not
    Dim aTxtDC          As Long     ' Edit's window DC - used for drawing in the non-client area.
    Dim pixAdj          As Long     ' number of pixels to adjust for the NC area.
    Dim lNumDigits      As Long     ' number of digits to display
    
    Dim aCtrl           As Control  ' generic object used to access the control's BorderStyle and Appearance properties.
    Dim aPtr            As Long     ' Pointer to the original control. Used in conjunction with aCtrl to access properties.
    Dim isRTB           As Boolean  ' Whether the control is an RTB or normal TextBox.
    
    If hWnd = 0 Then Exit Sub

    HasFocus = (GetFocus() = hWnd)

    ' Determine what kind of border this textbox has. This allows the routine to adjust
    '  for the border sizes. 3d = 2 pixels, fixed single = 1 pixel, none = 0 pixels. We'll
    '  use a hack to get an object reference directly to the control and access the
    '  properties directly. You could also use GetWindowLong and GWL_STYLE, but this is
    '  just as easy.
    aPtr = GetProp(hWnd, "ControlPtr")
    If aPtr <> 0 Then
        RtlMoveMemory aCtrl, aPtr, 4&
        isRTB = Not (TypeOf aCtrl Is TextBox)
        pixAdj = 0
        If aCtrl.BorderStyle = 1 Then pixAdj = pixAdj + 1
        If aCtrl.Appearance = 1 Then pixAdj = pixAdj + IIf(isRTB, 2, 1)
        RtlMoveMemory aCtrl, 0&, 4&
    Else
        pixAdj = 0
    End If

    aTxtDC = GetWindowDC(hWnd)
' -----------------------------------------
' PREPARE THE BACK BUFFER:
' -----------------------------------------

    ' Calculate the backbuffer's size
    aRECT.Right = GetProp(hWnd, "TrayWidth")
    GetWindowRect hWnd, rcWindow
    aRECT.Bottom = (rcWindow.Bottom - rcWindow.Top) - pixAdj
        
    ' Create our backbuffer
    backBuffDC = CreateCompatibleDC(aTxtDC)
    backBuffBmp = CreateCompatibleBitmap(aTxtDC, aRECT.Right, aRECT.Bottom)
    DeleteObject SelectObject(backBuffDC, backBuffBmp)
    ' Fill in the backbuffer with the correct background color
    aBrush = GetSysColorBrush(15)
    FillRect backBuffDC, aRECT, aBrush
    DeleteObject aBrush
    ' Make all of our text draw transparently
    SetBkMode backBuffDC, 1
        
' -----------------------------------------
' DETERMINE CURRENT SELECTION LINE NUMBERS:
' -----------------------------------------
    
    ' We first get the current selection's start and ending indices
    SendMessage hWnd, EM_GETSEL, curSel, curSelEnd
    ' Determine which lines the starting & ending indices are in
    curSelLine = SendMessage(hWnd, EM_LINEFROMCHAR, ByVal curSel, ByVal 0&) + 1
    curSelEndLine = SendMessage(hWnd, EM_LINEFROMCHAR, ByVal curSelEnd, ByVal 0&)
    
    If (isRTB And curSelEnd > SendMessage(hWnd, EM_LINEINDEX, ByVal curSelEndLine, ByVal 0&)) Or _
       (Not isRTB And curSelEnd >= SendMessage(hWnd, EM_LINEINDEX, ByVal curSelEndLine, ByVal 0&)) Then
        ' Past first character in the line. We want this line indicated as selected too
        curSelEndLine = curSelEndLine + 1
    End If
    
' -----------------------------------------------
' DETERMINE FONT TO USE FOR DRAWING LINE NUMBERS:
' -----------------------------------------------
    
    ' Create the font we're going to use to draw our numbers.
    If isRTB Then
        ' RTBs have a mix of fonts, so we therefore must CHOOSE a font to use for it. In this case,
        '  we'll use an 8 point Arial font.
        aFont = SelectObject(backBuffDC, CreateFont(-11, 0, 0, 0, 400, False, False, False, 0, 0, 0, 0, 0, "Arial"))
    Else
        ' Textboxes have a fixed height since the font is the font is the same for the entire
        '  control. As such, we'll use the same font as the control uses.
        aFont = SelectObject(backBuffDC, SendMessage(hWnd, WM_GETFONT, ByVal 0&, ByVal 0&))
    End If
    ' Set the font height in pixels for this specified font.
    GetTextExtentPoint32 backBuffDC, "qQ", 2, aPt
    FontHeight = aPt.Y
    
' -----------------------------------
' START THE LINE NUMBER DRAWING LOOP:
' -----------------------------------
    
    ' Get the top index
    lineTop = SendMessage(hWnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&) + 1
    If isRTB Then
        ' Find the last line using the x & y position at the bottom of the control
        aPt.x = 0
        aPt.Y = rcWindow.Bottom - rcWindow.Top - 2
        lineChar = SendMessage(hWnd, EM_CHARFROMPOS, ByVal 0&, aPt)
        lineBottom = SendMessage(hWnd, EM_LINEFROMCHAR, ByVal lineChar, ByVal 0&)
        ' Adjust the last line since all API values are index to 0 instead of 1
        If lineBottom <> 0 Then lineBottom = lineBottom + 1
    Else
        ' Calculate a theoretical last line for the edit based off the
        '  top line and the height of the edit
        If FontHeight <> 0 Then
            lineBottom = lineTop + ((rcWindow.Bottom - rcWindow.Top) \ FontHeight)
        End If
    End If
    ' Get the total number of lines in the edit.
    TotalLines = SendMessage(hWnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
    ' If the theoretical bottom is more than the total number of lines,
    '  cut it off at real # of total lines in the edit control.
    If lineBottom > TotalLines Or lineBottom = 0 Then lineBottom = TotalLines
    
    ' Loop thru and draw all of our line numbers.
    aRECT.Left = aRECT.Left + 2
    aRECT.Right = aRECT.Right - 3
    
    selRECT.Top = -1
    selRECT.Bottom = -1
    
    For curLine = lineTop To lineBottom
        If isRTB Then
            ' This may seem a little odd, but we're going to get the starting
            '  character index for each line, then translate it back to a x & y. This
            '  is because we're not guarenteed the same font height for each line.
            
            '  First we get the character index for this line
            lineChar = SendMessage(hWnd, EM_LINEINDEX, ByVal curLine - 1, ByVal 0&)
            ' Now we determine what the x & y coordinates for it is
            SendMessage hWnd, EM_POSFROMCHAR, aPt, ByVal lineChar
            ' Next we calculate the resulting RECT location for drawing in
            '  the line number based off that value.
            aRECT.Top = aPt.Y
            If curLine < lineBottom Then
                '  First we get the character index for this line
                lineChar = SendMessage(hWnd, EM_LINEINDEX, ByVal curLine, ByVal 0&)
                ' Now we determine what the x & y coordinates for it is
                SendMessage hWnd, EM_POSFROMCHAR, aPt, ByVal lineChar
                aRECT.Bottom = aPt.Y
            Else
                aRECT.Bottom = aRECT.Top + FontHeight
            End If
        Else
            ' We calculate the resulting RECT location for drawing in
            '  the line number based off the current line
            aRECT.Top = (curLine - lineTop) * FontHeight + 1
            aRECT.Bottom = aRECT.Top + FontHeight
        End If
        ' if it is one of our "Selected" lines, we want to do a little
        '  extra drawing to put the line number in "selected" colors.
        If (curSelLine <= curLine And curLine <= curSelEndLine) Or (curSelLine > curSelEndLine And curLine = curSelLine) Then
            ' Fill in the background for this line number with "selected" colors
            aBrush = CreateSolidBrush(MixColors(GetSysColor(IIf(HasFocus, 13, 15)), vbWhite, IIf(HasFocus, 50, 75)))
            FillRect backBuffDC, aRECT, aBrush
            DeleteObject aBrush
            If curLine = curSelLine Then selRECT.Top = aRECT.Top
            If curLine = curSelEndLine Or (curSelEndLine < curSelLine And curLine = curSelEndLine + 1) Then selRECT.Bottom = aRECT.Bottom
        End If
        SetTextColor backBuffDC, GetSysColor(8)
        ' Draw in the line number, right aligned and centered vertically
        lNumDigits = GetProp(hWnd, "NumDigits")
        If lNumDigits < 4 Then lNumDigits = 4
        DrawText backBuffDC, Format$(curLine, String$(lNumDigits, "0")), lNumDigits, aRECT, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    Next curLine
    
    selRECT.Left = aRECT.Left
    selRECT.Right = aRECT.Right
    
    ' Make sure first and last selected item hasn't scrolled out of visible area.
    If selRECT.Bottom = -1 And curSelEndLine > lineBottom Then selRECT.Bottom = (rcWindow.Bottom - rcWindow.Top) + 100
    If selRECT.Top = -1 And curSelLine < lineTop Then selRECT.Top = -100
    
    'Determine if we should draw the selection rectangle or not.
    If (curSelEndLine >= lineTop And curSelEndLine <= lineBottom) Or _
       (curSelLine >= lineTop And curSelLine <= lineBottom) Or _
       (curSelLine <= lineTop And curSelEndLine >= lineBottom) Then
        ' Draw our selection rectangle
        aBrush = SelectObject(backBuffDC, GetStockObject(5))
        aPen = SelectObject(backBuffDC, CreatePen(0, 1, GetSysColor(IIf(HasFocus, 13, 16))))
        Rectangle backBuffDC, selRECT.Left, selRECT.Top, selRECT.Right, selRECT.Bottom
        DeleteObject SelectObject(backBuffDC, aBrush)
        DeleteObject SelectObject(backBuffDC, aPen)
    End If
    
    ' Adjust our rectangle back to its original dimensions
    aRECT.Left = aRECT.Left - 2
    aRECT.Right = aRECT.Right + 3
        
' --------------------
' WRAP-UP AND CLEANUP:
' --------------------
    
    ' Draw a border between line numbers and the textbox
    If pixAdj > 0 Then
        aPen = SelectObject(backBuffDC, CreatePen(0, 1, GetSysColor(16)))
        MoveToEx backBuffDC, aRECT.Right - 1, 0, aPt
        LineTo backBuffDC, aRECT.Right - 1, (rcWindow.Bottom - rcWindow.Top) - pixAdj * 2
        DeleteObject SelectObject(backBuffDC, aPen)
    End If
    ' Transfer the backbuffered image to our control
    BitBlt aTxtDC, pixAdj, pixAdj, aRECT.Right, (rcWindow.Bottom - rcWindow.Top) - pixAdj * 2, backBuffDC, 0, 0, vbSrcCopy

    ' Clean up our graphics objects.
    If isRTB Then
        ' We created the font to do the numbering. We must therefore free that font handle by
        '  deleting it. We also replace the back buffer DC's original font so it can
        '  be desoyed with the DC.
        DeleteObject SelectObject(backBuffDC, aFont)
    Else
        ' We don't want to delete the control's font that we used to draw. Therefore
        '  we'll just release it back by replacing the back buffer's DC's original font.
        SelectObject backBuffDC, aFont
    End If
    DeleteDC backBuffDC
    DeleteObject backBuffBmp
    
    ' Fixed single borders on standard edits seem to behave a little odd. We have to manually redraw them, which is not
    '  necessary with the 3D style. We therefore need to manually draw a rectangle around the
    '  line number tray area and client area of the rectangle using the system "window frame" color (6).
    If pixAdj = 1 And Not isRTB Then
        ' Get the client dimensions
        GetClientRect hWnd, rcWindow
        ' Create a pen of the window frame (6) color
        aPen = SelectObject(aTxtDC, CreatePen(0, 1, GetSysColor(6)))
        ' Create a hollow (or null) brush so the rectangle is not "filled"
        aBrush = SelectObject(aTxtDC, GetStockObject(5))
        ' Draw the rectangle
        Rectangle aTxtDC, 0, 0, rcWindow.Right + GetProp(hWnd, "TrayWidth"), rcWindow.Bottom
        ' Clean up our temporary GDI32 objects.
        DeleteObject SelectObject(backBuffDC, aBrush)
        DeleteObject SelectObject(backBuffDC, aPen)
    End If
    
    ' Release the textbox's device context back to its window handle.
    ReleaseDC hWnd, aTxtDC
End Sub

Private Function MixColors(ByVal color1 As Long, ByVal color2 As Long, ByVal mixPercent As Long) As Long
    ' Generic function used for mixing two colors at a specific ratio.
    
    Dim red1 As Byte, blue1 As Byte, green1 As Byte
    Dim red2 As Byte, blue2 As Byte, green2 As Byte
    Dim red3 As Byte, blue3 As Byte, green3 As Byte
    
    Dim mixPer As Double
    
    mixPer = mixPercent / 100
    ' Get the component R/G/B values for our 2 colors
    red1 = color1 Mod 256
    green1 = (color1 \ 256) Mod 256
    blue1 = (color1 \ 65536)
    red2 = color2 Mod 256
    green2 = (color2 \ 256) Mod 256
    blue2 = (color2 \ 65536)
    ' Mix the two colors per our request
    red3 = red1 * mixPer + red2 * (1 - mixPer)
    green3 = green1 * mixPer + green2 * (1 - mixPer)
    blue3 = blue1 * mixPer + blue2 * (1 - mixPer)
    ' Convert it back into a long value and return it
    MixColors = RGB(red3, green3, blue3)
    
End Function
