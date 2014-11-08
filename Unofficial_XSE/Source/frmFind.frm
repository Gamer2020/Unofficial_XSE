VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Find and Replace"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   134
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   377
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "2000"
   Begin eXtremeScriptEditor.vcProgress vcProgress 
      Height          =   225
      Left            =   120
      Top             =   1620
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   397
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2520
      TabIndex        =   6
      Tag             =   "2007"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtReplace 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   555
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4080
      TabIndex        =   7
      Tag             =   "2008"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "Replace"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4080
      TabIndex        =   4
      Tag             =   "2004"
      Top             =   540
      Width           =   1455
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Match case"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Tag             =   "2006"
      Top             =   1020
      Width           =   3855
   End
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace All"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4080
      TabIndex        =   5
      Tag             =   "2005"
      Top             =   975
      Width           =   1455
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4080
      TabIndex        =   3
      Tag             =   "2002"
      Top             =   90
      Width           =   1455
   End
   Begin VB.TextBox txtFind 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   8
      X2              =   368
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Replace with"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Tag             =   "2003"
      Top             =   585
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find what"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Tag             =   "2001"
      Top             =   150
      Width           =   705
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lStart As Long
Private lOcc As Long
Private Break As Boolean
Private DontUpdateIndex As Boolean

Private Sub MsgBox2(sPrompt As String, Optional Buttons As VbMsgBoxStyle)
    
    ' When a MsgBox is displayed, the form
    ' lose focus and since when it's activated
    ' again the index gets updated we need to
    ' prevent that update
    DontUpdateIndex = True
    
    ' Display the MsgBox
    MsgBox sPrompt, Buttons
    
End Sub

Private Function InitializeSearch(Optional sText As String = vbNullString, Optional sFind As String = vbNullString) As Long
    
    ' If the length of the text is empty
    If LenB(sText) = 0 Then
        ' Use the current tab code
        sText = Document(frmMain.Tabs.SelectedTab).txtCode.text
    End If
    
    ' If the find text is empty
    If LenB(sFind) = 0 Then
        ' Use the current Find text
        sFind = txtFind.text
    End If

    If chkMatchCase.Value = vbUnchecked Then
        ' We aren't going to care about cases
        ' Set the return value to the number of occurences
        InitializeSearch = InStrCount(sText, sFind, , vbTextCompare)
    Else
        ' We are going to care about CaSES
        ' Set the return value to the number of occurences
        InitializeSearch = InStrCount(sText, sFind, , vbBinaryCompare)
    End If
    
    ' If everything went fine
    If InitializeSearch <> 0 Then
        ' Update the number of occurences
        lOcc = InitializeSearch
    End If
    
End Function

Private Sub SearchAndReplace(Optional MatchCase As Boolean = False, Optional ReplaceMode As Boolean = False, Optional ReplaceAll As Boolean = False)
Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Dim sText As String
Dim lLen As String
Dim lPos As Long
Dim sFind As String
Dim lFindLen As Long
Dim lReplaceLen As Long
    
    ' Set the find string
    sFind = txtFind.text
    
    ' If we don't care about matching cases
    If MatchCase = False Then
        ' Lowercase the find string
        sFind = LCase$(sFind)
    End If
    
    ' Retrieve the find and replace lenghts
    lFindLen = Len(sFind)
    lReplaceLen = Len(txtReplace.text)
    
    ' Set the mouse to busy
    MousePointer = vbHourglass

Begin:
    
    ' Calculate the length of the text
    lLen = SendMessage(Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, WM_GETTEXTLENGTH, 0&, ByVal 0&) + 1
    
    ' Allocate enough space
    sText = SysAllocStringLen(vbNullString, lLen)
    
    ' Get the text
    SendMessage Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, WM_GETTEXT, lLen, ByVal sText
    
    ' No matching?
    If MatchCase = False Then
        ' If so, lower case the text
        sText = LCase$(sText)
    End If
    
    ' Get the position of the first occurence
    lPos = InStr(lStart, sText, sFind)
    
    ' If we got something
    If lPos <> 0 Then
        
        ' Set the selection
        SendMessage Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_SETSEL, lPos - 1, ByVal (lPos - 1) + lFindLen
        
        If ReplaceMode = True Then
            
            ' If we are replacing, we need to replace
            ' the selection as well
            If LenB(txtReplace.text) <> 0 Then
                ' If the replace text is not empty
                SendMessage Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_REPLACESEL, 1&, ByVal txtReplace.text
            Else
                ' Else replace it with ""
                SendMessage Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_REPLACESEL, 1&, ByVal ""
            End If
            
        End If
        
        ' Make sure the caret is visible
        SendMessage Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_SCROLLCARET, 0&, ByVal 0&
        
        ' Adjust the start index
        lStart = lPos + 1
        
        If ReplaceMode = True Then
            
            ' If the replace length was higer than zero
            If lReplaceLen > 0 Then
                ' Increase the start index accordingly
                lStart = lStart + (lReplaceLen - lFindLen)
            End If
            
            If ReplaceAll = True Then
                
                ' If we're replacing everything, update the progress bar
                vcProgress.Value = vcProgress.Value + 1
                
                ' If the value is higher than then the max
                If vcProgress.Value >= vcProgress.Max Then
                    ' Hide it
                    vcProgress.Visible = False
                End If
                
                ' Determines whether there are mouse-button
                ' or keyboard messages
                If GetInputState <> 0 Then
                    
                    MyDoEvents
                   
                    If Break Then
                        ' User aborted
                        MsgBox2 LoadResString(2012), vbInformation
                        GoTo Finish
                    End If
                    
                End If
                
                ' Recursively replace
                GoTo Begin
                
            Else
                ' Single replace, hence select the next match
                SearchAndReplace MatchCase, False, False
                GoTo Finish
            End If
            
        End If
        
    Else
        
        If Not ReplaceAll Then
            
            ' If there are some other occurences
            If InitializeSearch(sText, sFind) <> 0 Then
                
                ' Reset the index
                lStart = 1
                MsgBox2 LoadResString(2010), vbInformation
                
                ' Search again
                SearchAndReplace MatchCase, False, False
                GoTo Finish
                
            Else
                ' Nothing more to search
                MsgBox2 LoadResString(2009), vbInformation
            End If
            
        Else
            
            If lOcc > 1 Then
                ' Two or more
                MsgBox2 LoadResString(2012) & vbSpace & Replace(LoadResString(2014), "x", CStr(lOcc)), vbInformation
            Else
                ' Only one occurence
                MsgBox2 LoadResString(2012) & vbSpace & LoadResString(2013), vbInformation
            End If
            
        End If
        
    End If
    
Finish:
    
    ' Reset the mouse
    MousePointer = vbDefault
    
End Sub

Private Sub cmdCancel_Click()
    
    ' This is needed to stop while
    ' doing a lengthy Replace operation
    Break = True
    
    ' Unload the form
    Unload Me
    
End Sub

Private Sub cmdFindNext_Click()
    
    ' Test if there's a list one occurence
    If InitializeSearch <> 0 Then
        ' If yes, search them
        SearchAndReplace CBool(chkMatchCase.Value)
    Else
        ' If not, tell the user
        MsgBox2 LoadResString(2011), vbInformation
    End If
    
End Sub

Private Sub cmdReplace_Click()
Dim lSelLen As Long
    
    ' Get the current selection length
    lSelLen = Document(frmMain.Tabs.SelectedTab).txtCode.SelLength
    
    ' If there is at least one occurence
    If InitializeSearch <> 0 Then
        
        ' Manually updated the current tab's Undo/Redo index
        Document(frmMain.Tabs.SelectedTab).StackIndex = Document(frmMain.Tabs.SelectedTab).StackIndex + 1
        
        ' Enable the Replacing flag
        Document(frmMain.Tabs.SelectedTab).IsReplacing = True
        
        ' If the current selection is not empty
        If lSelLen <> 0 Then
            
            ' Decrese the start index accordingly
            lStart = lStart - lSelLen
            
            ' If the start index goes too low
            If lStart < 1 Then
                ' Fix it
                lStart = 1
            End If
            
        End If
        
        ' Do the work
        SearchAndReplace CBool(chkMatchCase.Value), True
        
        ' Disable the Replacing flag
        Document(frmMain.Tabs.SelectedTab).IsReplacing = False
        
        ' Manually Update the stack content
        Document(frmMain.Tabs.SelectedTab).UpdateStack
        
    Else
        ' Nothing to replace
        MsgBox2 LoadResString(2011), vbInformation
    End If
    
End Sub

Private Sub cmdReplaceAll_Click()
Const ES_NOHIDESEL As Long = &H100
Dim lTemp As Long
    
    ' Store the number of occurences
    lTemp = InitializeSearch
    
    ' If there something that can be replaced
    If lTemp <> 0 Then
        
        ' Manually updated the current tab's Undo/Redo index
        Document(frmMain.Tabs.SelectedTab).StackIndex = Document(frmMain.Tabs.SelectedTab).StackIndex + 1
        
        ' Enable the Replacing flag
        Document(frmMain.Tabs.SelectedTab).IsReplacing = True
        
        ' Temporarily set the HideSelection property to True
        SetWindowLong Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, GWL_STYLE, GetWindowLong(Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, GWL_STYLE) Xor ES_NOHIDESEL
        
        ' We're replacing anything
        ' so the start index must be 1
        lStart = 1
        
        ' Update the progress bar value
        vcProgress.Value = 0
        vcProgress.Max = lTemp
        
        ' Make the progress bar visible
        vcProgress.Visible = True
        MyDoEvents
        
        ' Let's do the job
        SearchAndReplace CBool(chkMatchCase.Value), True, True
        
        ' Revert back the HideSelection property
        SetWindowLong Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, GWL_STYLE, GetWindowLong(Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, GWL_STYLE) Or ES_NOHIDESEL
        
        ' Hide the progress bar
        vcProgress.Visible = False
        
        ' Disable the Replacing flag
        Document(frmMain.Tabs.SelectedTab).IsReplacing = False
        
        ' Manually update the stack index
        Document(frmMain.Tabs.SelectedTab).UpdateStack
        
    Else
        ' Nothing to do
        MsgBox2 LoadResString(2011), vbInformation
    End If
    
End Sub

Private Sub cmdReset_Click()
    
    ' Reset anything
    txtFind.text = vbNullString
    txtReplace.text = vbNullString
    chkMatchCase.Value = vbUnchecked
    
    ' Recalculate the start index
    lStart = SendMessage(Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_LINEINDEX, -1, ByVal 0&)
    lStart = lStart + Document(frmMain.Tabs.SelectedTab).ActualColumn
    
    ' Set the focus to Find
    txtFind.SetFocus
    
    ' Disable the Reset
    cmdReset.Enabled = False
    
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    
    ' Check if the starting index
    ' should be updated
    If DontUpdateIndex Then
        ' If not, make sure it will
        ' be updated next time
        DontUpdateIndex = False
    Else
    
        ' Get the current line index
        lStart = SendMessage(Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_LINEINDEX, -1, ByVal 0&)
        
        ' Add the column value to it
        lStart = lStart + Document(frmMain.Tabs.SelectedTab).ActualColumn
        
    End If
    
End Sub

Private Sub Form_Load()
Dim lColumn As Long
Dim lSelLength As Long
Dim iLineLen As Integer
Dim sSelText As String

    Localize Me
    
    ' Get the selected text and selected length
    sSelText = GetTextBoxLine(Document(frmMain.Tabs.SelectedTab).txtCode.hWnd)
    lSelLength = Document(frmMain.Tabs.SelectedTab).txtCode.SelLength
    iLineLen = Len(sSelText)
    
    ' If something was selected
    If lSelLength <> 0 Then
    
        ' Get the column
        lColumn = Document(frmMain.Tabs.SelectedTab).ActualColumn
        
        ' If the length of selection is less or equal to the current line length
        If lSelLength <= iLineLen Then
            ' Adjust the selected text
            sSelText = Mid$(sSelText, lColumn - lSelLength, lSelLength)
        End If
        
        ' Update the Find textbox
        txtFind.text = sSelText
        txtFind.SelStart = Len(txtFind.text)
        
    End If
    
End Sub

Private Sub txtFind_Change()
    
    ' If the Find text is not empty
    If LenB(txtFind.text) <> 0 Then
        ' Enable the FindNext and the Reset
        cmdFindNext.Enabled = True
        cmdReset.Enabled = True
    Else
        ' Otherwhise disable FindNext
        cmdFindNext.Enabled = False
    End If
    
    ' Check if Replace/Replace All should
    ' be enabled or not
    ReplaceCheck
    
End Sub

Private Sub txtFind_KeyPress(KeyCode As Integer)
    
    ' If the user pressed Enter
    If KeyCode = vbKeyReturn Then
        ' Simulate a FindNext click
        KeyCode = 0
        cmdFindNext_Click
    End If
    
End Sub

Private Sub ReplaceCheck()
    
    ' Make sure the Replace text is not
    ' the same as the Find one
    If txtReplace.text <> txtFind.text Then
        ' If so, enabled Replace, Replace All
        ' and Reset
        cmdReplace.Enabled = True
        cmdReplaceAll.Enabled = True
        cmdReset.Enabled = True
    Else
        ' Otherwhise it doesn't make sense
        ' to have Replace enabled
        cmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
    End If
    
End Sub

Private Sub txtReplace_Change()
    ' Check wether Replace/Replace All
    ' need to be enabled or not
    ReplaceCheck
End Sub

