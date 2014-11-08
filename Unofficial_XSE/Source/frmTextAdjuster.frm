VERSION 5.00
Begin VB.Form frmTextAdjuster 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Text Adjuster"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTextAdjuster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "7000"
   Begin VB.TextBox txtMaxCount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4800
      MaxLength       =   2
      TabIndex        =   6
      Text            =   "34"
      Top             =   1965
      Width           =   495
   End
   Begin VB.PictureBox picLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   3915
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   225
      Width           =   15
      Begin VB.Line linLimit 
         BorderColor     =   &H00C8C8C8&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   76
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Enabled         =   0   'False
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Tag             =   "7001"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2760
      TabIndex        =   4
      Tag             =   "7003"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Enabled         =   0   'False
      Height          =   345
      Left            =   1440
      TabIndex        =   3
      Tag             =   "7002"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtSapp 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   5175
   End
   Begin VB.TextBox txtCharCount 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1965
      Width           =   495
   End
   Begin VB.TextBox txtToAdjust 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      MaxLength       =   999
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      Height          =   195
      Left            =   4665
      TabIndex        =   7
      Top             =   1995
      Width           =   60
   End
   Begin VB.Menu mnuCustomPopup 
      Caption         =   "CustomMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Enabled         =   0   'False
         HelpContextID   =   7004
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
         HelpContextID   =   7005
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
         HelpContextID   =   7006
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
         HelpContextID   =   7007
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
         Enabled         =   0   'False
         HelpContextID   =   7008
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         HelpContextID   =   20
      End
      Begin VB.Menu mnuInsertSpecialChar 
         Caption         =   "Insert Special Char"
         HelpContextID   =   7009
      End
   End
End
Attribute VB_Name = "frmTextAdjuster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lMaxCount As Long
Private PastedText As Boolean
Private IsAdjusting As Boolean
Private sBrackets() As String

Private Const WM_KILLFOCUS = &H8&
Private Const WM_CONTEXTMENU = &H7B&

Private cSubclasser As cSelfSubclasser

Public Sub GetLineLen()
    txtCharCount.text = Len(GetTextBoxLine(txtToAdjust.hWnd, -1))
End Sub

Private Sub cmdClear_Click()

    If MsgBox(LoadResString(7010), vbYesNo + vbExclamation) = vbNo Then
        Exit Sub
    End If
    
    txtToAdjust.text = vbNullString
    txtSapp.text = vbNullString
    cmdClear.Enabled = False
    
    txtToAdjust.SetFocus
    
End Sub

Private Sub ReAdjust()
    
    If InStrB(txtToAdjust.text, vbSpace) Then
        txtToAdjust.text = Replace(txtToAdjust.text, vbNewLine, vbSpace)
    Else
        txtToAdjust.text = Replace(txtToAdjust.text, vbNewLine, vbNullString)
    End If
    
    If IsWrapped = False Then
        txtToAdjust_Change
    End If
    
End Sub

Private Sub TrimText()
Dim sArray() As String
Dim i As Long
        
    SplitB txtToAdjust.text, sArray(), vbNewLine

    For i = LBound(sArray) To UBound(sArray)
        sArray(i) = Trim$(Left$(sArray(i), lMaxCount))
    Next i

    txtToAdjust.text = Join(sArray(), vbNewLine)

End Sub

Private Sub cmdConvert_Click()
Const sNewLine As String = "\n"
Const sNewParagraph As String = "\l"
Const sNewPage As String = "\p"
Dim sArray() As String, sArray2() As String
Dim i As Long
Dim j As Long
Dim sTemp As String
Dim cString As cStringBuilder

    TrimText
    
    Set cString = New cStringBuilder
    txtSapp.text = vbNullString
    
    sTemp = txtToAdjust.text
    
    PutBrackets sTemp
    
    SplitB sTemp, sArray(), vbNewLine & vbNewLine
    
    cString.Append "= "

    For i = LBound(sArray) To UBound(sArray)
        If InStrB(1, sArray(i), vbNewLine, vbBinaryCompare) <> 0 Then
            SplitB sArray(i), sArray2(), vbNewLine
            If UBound(sArray2) = 1 Then
                cString.Append sArray2(LBound(sArray2)) & sNewLine
                cString.Append sArray2(UBound(sArray2)) & sNewPage
            Else
                cString.Append sArray2(LBound(sArray2)) & sNewLine
                For j = LBound(sArray2) + 1 To UBound(sArray2) - 1
                    cString.Append sArray2(j) & sNewParagraph
                Next j
                cString.Append sArray2(UBound(sArray2)) & sNewPage
            End If
        Else
            cString.Append sArray(i) & sNewPage
        End If
    Next i
    
    Erase sArray
    Erase sArray2
    
    cString.Remove cString.Length - 2, 2
    
    If cString.Length > 1024 Then
        cString.Remove 1024, cString.Length - 1024
    End If
    
    txtSapp.text = cString.ToString
        
    txtSapp.SelStart = 0
    txtSapp.SelLength = Len(txtSapp.text)
    txtSapp.SetFocus
    
    SafeClipboardSet txtSapp.text
    
End Sub

Private Sub cmdInsert_Click()
    'Document(Tabs.SelectedTab).txtCode.SelText = txtSapp.text
    SendMessageStr Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_REPLACESEL, 1&, txtSapp.text
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    
    Select Case KeyCode
    
        Case vbKeyEscape
            
            Unload Me
            
        Case vbKeyH - 64, vbKeyI - 64, vbKeyJ - 64, vbKeyM - 64
            
            If GetKeyState(vbKeyControl) < 0 Then
                KeyCode = 0
            End If
            
    End Select
    
End Sub

Private Sub Form_Load()
    
    Localize Me
    
    lMaxCount = ReadIniString(App.Path & IniFile, "TextAdjuster", "TextLimit", 34)
    txtMaxCount.text = lMaxCount
    picLine.Left = lMaxCount * 7 + 16
    
    Set cSubclasser = New cSelfSubclasser
    
    If cSubclasser.ssc_Subclass(txtToAdjust.hWnd, , 1, Me) = True Then
        cSubclasser.ssc_AddMsg txtToAdjust.hWnd, eMsgWhen.MSG_BEFORE, WM_CONTEXTMENU
    End If
    
    ReDim sBrackets(0) As String
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WriteStringToIni App.Path & IniFile, "TextAdjuster", "TextLimit", lMaxCount
    Set cSubclasser = Nothing
End Sub

Private Sub mnuClear_Click()
    cmdClear_Click
End Sub

Private Sub mnuCopy_Click()
    SendMessage txtToAdjust.hWnd, WM_COPY, 0, ByVal 0&
End Sub

Private Sub mnuCut_Click()
    SendMessage txtToAdjust.hWnd, WM_CUT, 0, ByVal 0&
End Sub

Private Sub mnuInsertSpecialChar_Click()
    Show2 frmInsertChar, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuPaste_Click()
    PastedText = True
    SendMessage txtToAdjust.hWnd, WM_PASTE, 0, ByVal 0&
End Sub

Private Sub mnuSelectAll_Click()
    SendMessage txtToAdjust.hWnd, EM_SETSEL, 0&, ByVal -1
    SendMessage txtToAdjust.hWnd, EM_SCROLLCARET, 0&, ByVal 0&
End Sub

Private Sub mnuUndo_Click()
Const EM_UNDO = &HC7
    SendMessage txtToAdjust.hWnd, EM_UNDO, 0, ByVal 0&
End Sub

Private Sub txtCharCount_GotFocus()
    SendMessage txtCharCount.hWnd, WM_KILLFOCUS, 0&, ByVal 0&
End Sub

Private Sub txtCharCount_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    txtCharCount.SelLength = 0
End Sub

Private Sub txtMaxCount_LostFocus()
    If LenB(txtMaxCount.text) = 0 Then
        txtMaxCount.text = lMaxCount
    Else
        If CByte(txtMaxCount.text) < 10 Then
            lMaxCount = 10
            txtMaxCount = lMaxCount
        End If
    End If
End Sub

Private Sub txtSapp_Change()
    If LenB(txtSapp.text) <> 0 Then
        cmdInsert.Enabled = True
    Else
        cmdInsert.Enabled = False
    End If
End Sub

Private Function IsWrapped() As Boolean
Dim sArray() As String
Dim i As Long

    IsWrapped = True

    SplitB txtToAdjust.text, sArray(), vbNewLine
    
    For i = LBound(sArray) To UBound(sArray)
        If Len(sArray(i)) > lMaxCount Then
            IsWrapped = False
            Exit For
        End If
    Next i
    
End Function

Private Sub GetBrackets()
Dim sText As String
Dim lCount As Long
Dim lStart As Long
Dim lStart2 As Long
Dim lSelStart As Long

    sText = txtToAdjust.text
    
    ' Search first open bracket
    lStart = InStr(sText, "[")
    
    If lStart <> 0 Then
        
        lSelStart = txtToAdjust.SelStart
        
        If LenB(sBrackets(0)) <> 0 Then
            lCount = UBound(sBrackets) + 1
        End If
        
        Do While lStart <> 0
            
            lStart2 = InStr(lStart + 1, sText, "]")
            
            If lStart2 <> 0 Then
            
                If lCount >= UBound(sBrackets) Then
                    ReDim Preserve sBrackets(lCount + 10) As String
                End If
            
                If lStart2 - lStart > 1 Then
                    
                    sBrackets(lCount) = Mid$(sText, lStart + 1, lStart2 - (lStart + 1))
                    lCount = lCount + 1
                    
                    sText = Left$(sText, lStart) & Mid$(sText, lStart2)
                    
                End If
                
                lStart = InStr(lStart + 1, sText, "[")
                
            Else
                Exit Do
            End If
            
        Loop
        
        If lCount > 0 Then
            ReDim Preserve sBrackets(lCount - 1) As String
            txtToAdjust.text = sText
            txtToAdjust.SelStart = lSelStart
        End If
        
    End If

End Sub

Private Sub PutBrackets(ByRef sText As String)
Dim lCount As Long
Dim lStart As Long
Dim lStart2 As Long
    
    ' Search first open bracket
    lStart = InStr(sText, "[")
    
    If lStart <> 0 Then
        
        Do While lStart <> 0
            
            lStart2 = InStr(lStart + 1, sText, "]")
            
            If lStart2 <> 0 Then
            
                If lStart2 - lStart = 1 Then
                    
                    If lCount <= UBound(sBrackets) Then
                        sText = (Left$(sText, lStart) & sBrackets(lCount)) & Mid$(sText, lStart2)
                        lCount = lCount + 1
                    Else
                        Exit Do
                    End If

                End If
                
                lStart = InStr(lStart + 1, sText, "[")
                
            Else
                Exit Do
            End If
            
        Loop
        
    End If

End Sub

Private Sub txtToAdjust_Change()
Dim lLineLen As Long
Dim lLineIndex As Long
Dim lCurrentLine As Long
Dim lSpaceIndex As Long
Dim lNewSelStart As Long
Dim sTemp As String
Dim sArray() As String
Dim i As Long

    If LenB(txtToAdjust.text) <> 0 Then
        cmdConvert.Enabled = True
        cmdClear.Enabled = True
    Else
        cmdConvert.Enabled = False
        cmdClear.Enabled = False
        txtCharCount.text = 0
        Exit Sub
    End If

    If IsAdjusting Then Exit Sub
    IsAdjusting = True
    
    GetBrackets
    
    If IsWrapped = False Then
        
        sTemp = txtToAdjust.text
        txtToAdjust.MaxLength = 0
        LockUpdate txtToAdjust.hWnd
        
        For i = 1 To 40
            
            For lCurrentLine = 0 To SendMessage(txtToAdjust.hWnd, EM_GETLINECOUNT, 0&, ByVal 0&) - 1
                GoSub Adjust
            Next lCurrentLine
            
            If IsWrapped = True Then Exit For
            
        Next i
        
        txtToAdjust.MaxLength = 999
        txtToAdjust.SelLength = 0
        
        If Not PastedText Then
            txtToAdjust.SelStart = lNewSelStart
        Else
            txtToAdjust.SelStart = Len(txtToAdjust.text)
            PastedText = False
        End If
        
        UnlockUpdate txtToAdjust.hWnd
        SendMessage txtToAdjust.hWnd, EM_SCROLLCARET, &H0, ByVal &H0
        
    End If
    
    IsAdjusting = False
    Exit Sub
    
Adjust:
    
    SplitB sTemp, sArray(), vbNewLine

    lLineLen = Len(sArray(lCurrentLine))
    lLineIndex = SendMessage(txtToAdjust.hWnd, EM_LINEINDEX, lCurrentLine, ByVal 0&)

    If lLineLen > lMaxCount Then
        
        lSpaceIndex = InStrRev(Mid$(GetTextBoxLine(txtToAdjust.hWnd, lCurrentLine), 1, lMaxCount), vbSpace)
        
        If lSpaceIndex <> 0 Then
            
            If lSpaceIndex <= lMaxCount Then
                sTemp = RTrim$(Mid$(sTemp, 1, lLineIndex + lSpaceIndex - 1)) & vbNewLine & Mid$(sTemp, lLineIndex + lSpaceIndex + 1)
            Else
                sTemp = RTrim$(Mid$(sTemp, 1, lLineIndex + lMaxCount)) & vbNewLine & Mid$(sTemp, lLineIndex + lMaxCount + 1)
            End If
            
        Else
            sTemp = Mid$(sTemp, 1, lLineIndex + lMaxCount) & vbNewLine & Mid$(sTemp, lLineIndex + 1 + lMaxCount)
        End If
        
        lNewSelStart = txtToAdjust.SelStart
        
        If (lNewSelStart - lLineIndex) > lMaxCount Then
            lNewSelStart = lNewSelStart + 2
        End If
        
        txtToAdjust.text = sTemp
    
    End If

    Return

End Sub

Private Sub txtCharCount_Change()
    If CByte(txtCharCount.text) < Int((lMaxCount / 100) * 50) Then
        txtCharCount.ForeColor = &HC000&
    ElseIf CByte(txtCharCount.text) >= Int((lMaxCount / 100) * 50) And CByte(txtCharCount.text) < Int((lMaxCount / 100) * 75) Then
        txtCharCount.ForeColor = &H80FF&
    ElseIf CByte(txtCharCount.text) >= Int((lMaxCount / 100) * 75) Then
        txtCharCount.ForeColor = vbRed
    End If
End Sub

Private Sub txtMaxCount_Change()
    If LenB(txtMaxCount.text) <> 0 Then
        If IsNumeric(txtMaxCount.text) Then
            If CByte(txtMaxCount.text) <= 40 Then
                If CByte(txtMaxCount.text) >= 10 Then
                    lMaxCount = CByte(txtMaxCount.text)
                End If
            ElseIf CByte(txtMaxCount.text) > 40 Then
                lMaxCount = 40
                txtMaxCount.text = lMaxCount
                txtMaxCount.SelStart = Len(txtMaxCount.text)
            End If
            If Len(txtMaxCount.text) = 2 Then
                picLine.Left = lMaxCount * 7 + 16
                'TrimText
                ReAdjust
                GetLineLen
            End If
        Else
            txtMaxCount.text = lMaxCount
            txtMaxCount.SelStart = Len(txtMaxCount.text)
        End If
    End If
End Sub

Private Sub txtToAdjust_Click()
    GetLineLen
End Sub

Private Sub txtToAdjust_KeyDown(KeyCode As Integer, Shift As Integer)
    GetLineLen
End Sub

Private Sub txtToAdjust_KeyPress(KeyCode As Integer)
Dim sTemp As String
    
    Select Case KeyCode
        
        Case vbKeySpace
            
            If CByte(txtCharCount.text) >= lMaxCount Then
                KeyCode = vbKeyReturn
            End If
            
        Case vbKeyReturn, vbKeyBack, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight
        
        Case vbKeyV - 64 'Ctrl+V
            
            KeyCode = 0
            sTemp = Clipboard.GetText
            
            If LenB(sTemp) <> 0 Then 'And Len(sTemp) <= (lMaxCount - CByte(txtCharCount.text)) Then
                mnuPaste_Click
            End If

        Case Else
            
            If CByte(txtCharCount.text) > lMaxCount Then
                KeyCode = 0
                txtToAdjust_Change
            End If
            
    End Select
    
End Sub

Private Sub txtToAdjust_KeyUp(KeyCode As Integer, Shift As Integer)
    
    GetLineLen
    
    If KeyCode = vbKeyE Then
        If Shift = vbCtrlMask Then
            SendMessageStr txtToAdjust.hWnd, EM_REPLACESEL, 1&, "é"
        End If
    End If
    
End Sub

Private Sub ShowCustomMenu()
Const EM_CANUNDO = &HC6
Dim tmpString As String
    
    txtToAdjust.Enabled = False
    txtToAdjust.Enabled = True
    txtToAdjust.SetFocus

    If LenB(txtToAdjust.text) <> 0 Then
        mnuCopy.Enabled = True
        mnuCut.Enabled = True
    Else
        mnuCopy.Enabled = False
        mnuCut.Enabled = False
    End If

    mnuUndo.Enabled = SendMessage(txtToAdjust.hWnd, EM_CANUNDO, 0, ByVal 0&)
    mnuClear.Enabled = cmdClear.Enabled

    tmpString = Clipboard.GetText
    
    If LenB(tmpString) <> 0 Then
        mnuPaste.Enabled = True
    Else
        mnuPaste.Enabled = False
    End If
    
    'If CByte(txtCharCount.text) < lMaxCount Then
    '    mnuInsertSpecialChar.Enabled = True
    'Else
    '    mnuInsertSpecialChar.Enabled = False
    'End If

    PopupMenu mnuCustomPopup
        
End Sub

Private Sub txtToAdjust_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    GetLineLen
End Sub

Private Sub txtToAdjust_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    GetLineLen
End Sub

'- ordinal #1
Private Sub myWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
        
        Select Case uMsg
        
            Case WM_CONTEXTMENU
                
                ShowCustomMenu
                
                bHandled = True
                lReturn = 1
            
        End Select

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
' *************************************************************
        
End Sub
