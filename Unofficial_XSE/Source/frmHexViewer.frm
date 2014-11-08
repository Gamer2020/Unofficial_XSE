VERSION 5.00
Begin VB.Form frmHexViewer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Hex Viewer"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "frmHexViewer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "14000"
   Begin VB.Timer tmrWatch 
      Interval        =   800
      Left            =   2880
      Top             =   0
   End
   Begin VB.TextBox txtPrefix 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "0x"
      Top             =   120
      Width           =   255
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1815
      LargeChange     =   96
      Left            =   6240
      SmallChange     =   12
      TabIndex        =   11
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox txtStatus 
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
      Top             =   2370
      Width           =   6375
   End
   Begin VB.CheckBox chkAscii 
      Caption         =   "ASCII/Poké"
      Height          =   255
      Left            =   3420
      TabIndex        =   4
      Top             =   165
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrevLarge 
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   5
      Top             =   150
      Width           =   330
   End
   Begin VB.CommandButton cmdPrevSmall 
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   6
      Top             =   150
      Width           =   330
   End
   Begin VB.CommandButton cmdNextSmall 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   7
      Top             =   150
      Width           =   330
   End
   Begin VB.CommandButton cmdNextLarge 
      Caption         =   ">>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5760
      TabIndex        =   8
      Top             =   150
      Width           =   330
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   4620
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   10
      Top             =   585
      Width           =   15
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   107
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1605
      Left            =   1260
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   9
      Top             =   585
      Width           =   15
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   107
      End
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Enabled         =   0   'False
      Height          =   345
      Left            =   1320
      TabIndex        =   3
      Tag             =   "14001"
      Top             =   75
      Width           =   855
   End
   Begin VB.TextBox txtHex 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   6015
   End
   Begin VB.TextBox txtOffset 
      Height          =   285
      Left            =   360
      MaxLength       =   7
      TabIndex        =   2
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "frmHexViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bBytes(95) As Byte
Private lFileLen As Long
Private lOldChecksum As Long
Private lOldOffset As Long
Private sFile As String

Private WithEvents vsb As CLongScroll
Attribute vsb.VB_VarHelpID = -1
Private hexOffset As clsHexBox
Attribute hexOffset.VB_VarHelpID = -1

Private Sub chkAscii_Click()
    GetHex
End Sub

Private Function GetChecksum() As Long
Dim iFileNum As Integer
Dim bTempBytes(95) As Byte
Dim i As Byte

    iFileNum = FreeFile
    
    ' Open the ROM and get a chunk of byte
    Open sFile For Binary As #iFileNum
        Get #iFileNum, CLng("&H" & txtOffset.text) + 1, bTempBytes
    Close #iFileNum
    
    ' Calculate the sum of the values
    For i = LBound(bTempBytes) To UBound(bTempBytes)
        GetChecksum = GetChecksum + bTempBytes(i)
    Next i

End Function

Private Sub GetHex(Optional Refresh As Boolean = False)
Const sHexPrefix As String = "0x"
Const vbDot As String = "."
Dim iFileNum As Integer
Dim cString As cStringBuilder
Dim i As Byte
Dim j As Byte
Dim bUBound As Byte
Dim bTemp(0) As Byte
    
    ' Store the current loaded file
    sFile = Document(frmMain.Tabs.SelectedTab).LoadedFile
    
    ' If the file is empty, exit
    If LenB(sFile) = 0 Then
        Exit Sub
    End If
    
    iFileNum = FreeFile
    Set cString = New cStringBuilder
    
    ' Open the file
    Open sFile For Binary As #iFileNum
        
        ' Get the bytes
        Get #iFileNum, CLng("&H" & txtOffset.text) + 1, bBytes
        
        ' Store the file length
        lFileLen = LOF(iFileNum)
        
        ' If the length is less or equal to 16 MB
        If lFileLen <= &H1000000 Then
            ' Offset can't use more than 6 chars
            txtOffset.MaxLength = 6
            txtOffset.text = Left$(txtOffset.text, 6)
        Else
            ' Otherwhise we may need 7 chars
            txtOffset.MaxLength = 7
        End If
        
    Close #iFileNum
    
    ' Get the upper bound
    bUBound = UBound(bBytes) \ 8 + 1
    
    ' 8 lines
    For i = 0 To 7
        
        ' Add the offset
        cString.Append sHexPrefix & PadHex$(CLng("&H" & txtOffset.text) + (i * bUBound), 7) & Space$(3)
        
        For j = i * bUBound To i * bUBound + bUBound - 1
            
            ' Add the byte
            cString.Append PadHex$(bBytes(j), 2)
            
            ' After each 2 bytes, add a space
            If (j + 1) Mod 2 = 0 Then cString.Append Space$(1)
            
        Next j
        
        cString.Append Space$(2)
        
        For j = i * bUBound To i * bUBound + bUBound - 1
            
            ' ASCII mode
            If chkAscii.Value = vbUnchecked Then
                
                ' Add a character/dot, depending on its value
                If bBytes(j) > 31 And bBytes(j) < 127 Then
                    cString.Append ChrW$(bBytes(j))
                Else
                    cString.Append vbDot
                End If
                
            ' Poké mode
            Else
                
                ' Store the current byte
                bTemp(0) = bBytes(j)
                
                ' If the converted string has a single char, add it
                If Len(Sapp2Asc(bTemp)) = 1 Then
                    cString.Append Sapp2Asc(bTemp)
                Else
                    ' More than one char, add a dot instead
                    cString.Append vbDot
                End If
                
            End If
            
        Next j
        
        ' If we didn't reach the last line, add vbNewLine
        If i < 7 Then
            cString.Append vbNewLine
        End If
        
    Next i
    
    ' Check which navigation buttons should be enabled/dsaibled
    CheckNavigation (txtOffset.text)
    
    ' Set the text
    txtHex.text = cString.ToString
    
    ' If the form is the active window
    If GetActiveWindow = Me.hWnd Then
        ' If the Refresh flag is not set
        If Refresh = False Then
            ' Set the caret position
            txtHex.SelStart = bUBound
            txtHex.SelLength = 0
            txtHex.SetFocus
        End If
    End If
    
    ' Update the status text
    UpdateStatus
    
    ' Save the current offset and checksum
    lOldOffset = CLng("&H" & txtOffset.text)
    lOldChecksum = GetChecksum
    
End Sub

Public Sub cmdGo_Click()
    GetHex
    vsb.Value = CLng("&H" & txtOffset.text)
End Sub

Private Sub NavigateOffset(lMove As Long)
    If LenB(txtOffset.text) <> 0 Then
        If Len(txtOffset.text) >= 6 Then
            If IsHex(txtOffset.text) Then
                txtOffset.text = PadHex$(CLng("&H" & txtOffset.text) + lMove, txtOffset.MaxLength)
                cmdGo_Click
            End If
        End If
    End If
End Sub

Private Sub cmdNextLarge_Click()
    NavigateOffset ((UBound(bBytes) + 1) \ 8) * 7
    vsb.Value = CLng("&H" & txtOffset.text)
End Sub

Private Sub cmdNextSmall_Click()
    NavigateOffset (UBound(bBytes) + 1) \ 8
    vsb.Value = CLng("&H" & txtOffset.text)
End Sub

Private Sub cmdPrevLarge_Click()
    NavigateOffset -((UBound(bBytes) + 1) \ 8) * 7
    vsb.Value = CLng("&H" & txtOffset.text)
End Sub

Private Sub cmdPrevSmall_Click()
    NavigateOffset -(UBound(bBytes) + 1) \ 8
    vsb.Value = CLng("&H" & txtOffset.text)
End Sub

Private Sub Form_Activate()

    If ActiveControl.name = "txtStatus" Then
        txtHex.SelStart = (UBound(bBytes) + 1) \ 8
        txtHex.SelLength = 0
        txtHex.SetFocus
    End If
    
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()

    Localize Me
    
    Set hexOffset = New clsHexBox
    Set hexOffset.TextBox = txtOffset
    
    sFile = Document(frmMain.Tabs.SelectedTab).LoadedFile
    txtOffset.text = Document(frmMain.Tabs.SelectedTab).txtOffset.text
    
    If LenB(txtOffset.text) = 0 Then
        txtOffset.text = PadHex$(0, 6)
    End If
    
    Set vsb = New CLongScroll
    Set vsb.Client = VScroll
    
    chkAscii.Value = ReadIniString(App.Path & IniFile, "Options", "HexViewerFilter", 0)
    
    GetHex
    txtHex.SelStart = (UBound(bBytes) + 1) \ 8
    txtHex.SelLength = 0
        
    vsb.Min = 0
    vsb.Max = lFileLen - (UBound(bBytes) + 1)
    vsb.SmallChange = (UBound(bBytes) + 1) \ 8
    vsb.LargeChange = (UBound(bBytes) + 1)
    
    vsb.Value = CLng("&H" & txtOffset.text)
    GetHex True
    
End Sub

Private Sub CheckNavigation(sOffset As String)
Dim lOffset As Long
Dim bUBound As Byte

    lOffset = CLng("&H" & sOffset)
    bUBound = UBound(bBytes) + 1
    
    If lOffset <= (bUBound \ 8) - 1 Then
        cmdPrevSmall.Enabled = False
        cmdPrevLarge.Enabled = False
        cmdNextSmall.Enabled = True
        cmdNextLarge.Enabled = True
    ElseIf lOffset > (bUBound \ 8) - 1 And lOffset < (bUBound \ 8) * 7 Then
        cmdPrevSmall.Enabled = True
        cmdPrevLarge.Enabled = False
        cmdNextSmall.Enabled = True
        cmdNextLarge.Enabled = True
    ElseIf lOffset >= (bUBound \ 8) * 7 And lOffset < lFileLen - bUBound Then
        cmdPrevSmall.Enabled = True
        cmdPrevLarge.Enabled = True
        cmdNextSmall.Enabled = True
        If lOffset <= lFileLen - bUBound * 2 + bUBound \ 2 Then
            cmdNextLarge.Enabled = True
        Else
            cmdNextLarge.Enabled = False
        End If
    ElseIf lOffset >= lFileLen - bUBound Then
        cmdPrevSmall.Enabled = True
        cmdPrevLarge.Enabled = True
        cmdNextSmall.Enabled = False
        cmdNextLarge.Enabled = False
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    WriteStringToIni App.Path & IniFile, "Options", "HexViewerFilter", chkAscii.Value
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set vsb = Nothing
    
End Sub

Private Sub tmrWatch_Timer()
Dim lNewChecksum As Long

    sFile = Document(frmMain.Tabs.SelectedTab).LoadedFile

    If LenB(sFile) = 0 Then
        Exit Sub
    ElseIf GetExt(sFile) <> "gba" Then
        Exit Sub
    End If

    If LenB(txtOffset.text) <> 0 Then
        If Len(txtOffset.text) >= 6 And IsHex(txtOffset.text) Then
            If lOldOffset = CLng("&H" & txtOffset.text) Then
                lNewChecksum = GetChecksum
                If lNewChecksum <> lOldChecksum Then
                    GetHex (True)
                End If
            End If
        End If
    End If

End Sub

Public Sub ToggleEnable(sFile As String)
    
    If GetExt(sFile) = "gba" Then
        cmdGo.Enabled = True
        txtHex.Enabled = True
        txtOffset.Enabled = True
        chkAscii.Enabled = True
        VScroll.Enabled = True
        GetHex (True)
    Else
        cmdGo.Enabled = False
        cmdPrevLarge.Enabled = False
        cmdPrevSmall.Enabled = False
        cmdNextSmall.Enabled = False
        cmdNextLarge.Enabled = False
        txtOffset.Enabled = False
        chkAscii.Enabled = False
        txtHex.Enabled = False
        VScroll.Enabled = False
    End If
    
End Sub

Private Sub UpdateStatus()
Dim iSelStart As Integer
Dim bLine As Byte
Dim iIndex As Integer
Dim bTempByte As Byte
Dim bCurrentByte As Byte
Dim bUBound As Byte
Dim lOffset As Long

    iSelStart = txtHex.SelStart
    bLine = Int(iSelStart / 58)
    iIndex = iSelStart - (bLine * 58)
    
    If iIndex > 11 And iIndex < 41 Then
        If AscW(Mid$(txtHex.text, iSelStart + 1, 1)) <> 32 Then
            Select Case iIndex
                Case 12, 13
                    bTempByte = 0
                Case 14, 15
                    bTempByte = 1
                Case 17, 18
                    bTempByte = 2
                Case 19, 20
                    bTempByte = 3
                Case 22, 23
                    bTempByte = 4
                Case 24, 25
                    bTempByte = 5
                Case 27, 28
                    bTempByte = 6
                Case 29, 30
                    bTempByte = 7
                Case 32, 33
                    bTempByte = 8
                Case 34, 35
                    bTempByte = 9
                Case 37, 38
                    bTempByte = 10
                Case 39, 40
                    bTempByte = 11
            End Select
        Else
            Exit Sub
        End If
    
    ElseIf iIndex > 43 And iIndex < 56 Then
        bTempByte = iIndex - 44
    Else
        Exit Sub
    End If
    
    bUBound = (UBound(bBytes) + 1) \ 8
    bCurrentByte = bTempByte + bUBound * bLine
    
    If LenB(txtOffset.text) <> 0 Then
        If IsHex(txtOffset.text) Then
            lOffset = CLng("&H" & txtOffset.text) + (bLine * bUBound) + bTempByte
            txtStatus.text = LoadResString(13038) & ": 0x" & PadHex$(lOffset, 7) & " (" & lOffset & ") | " & LoadResString(14004) & ": 0x" & PadHex$(bBytes(bCurrentByte), 2) & " (" & bBytes(bCurrentByte) & ")"
        End If
    End If
    
End Sub

Private Sub txtHex_Click()
    UpdateStatus
End Sub

Private Sub txtHex_KeyDown(KeyCode As Integer, Shift As Integer)
    UpdateStatus
End Sub

Private Sub txtHex_KeyUp(KeyCode As Integer, Shift As Integer)
    UpdateStatus
End Sub

Private Sub txtHex_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    UpdateStatus
End Sub

Private Sub txtHex_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    UpdateStatus
End Sub

Private Sub txtHex_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    UpdateStatus
End Sub

Private Sub txtOffset_Change()
    
    lFileLen = FileLength(sFile)
    
    If lFileLen <= &H1000000 Then
        txtOffset.MaxLength = 6
        txtOffset.text = Left$(txtOffset.text, 6)
    Else
        txtOffset.MaxLength = 7
    End If
    
    If LenB(txtOffset.text) <> 0 Then
    
        If Len(txtOffset.text) >= 6 And IsHex(txtOffset.text) Then
            
            If lFileLen > 0 Then
                If CLng("&H" & txtOffset.text) > lFileLen - UBound(bBytes) Then
                    txtOffset.text = PadHex$(lFileLen - (UBound(bBytes) + 1), txtOffset.MaxLength)
                End If
            End If
            
            cmdGo.Enabled = True
            CheckNavigation (txtOffset.text)
            
        Else
            cmdGo.Enabled = False
        End If
        
    Else
       cmdGo.Enabled = False
    End If
End Sub

Private Sub txtOffset_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If LenB(txtOffset.text) <> 0 Then
            If IsHex(txtOffset.text) Then
                KeyCode = 0
                cmdGo_Click
            End If
        End If
    End If
End Sub

Private Sub vsb_Change()
    txtOffset.text = PadHex$(vsb.Value, txtOffset.MaxLength)
    GetHex
End Sub

Private Sub vsb_Scroll()
    txtOffset.text = PadHex$(vsb.Value, txtOffset.MaxLength)
    GetHex
End Sub
