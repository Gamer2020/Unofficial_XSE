VERSION 5.00
Begin VB.Form frmBatch 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Batch Compiler"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBatch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   241
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "5000"
   Begin VB.CheckBox chkSelect 
      Caption         =   "Select All"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Tag             =   "5013"
      Top             =   3150
      Width           =   2055
   End
   Begin VB.ListBox lstHeaders 
      Height          =   510
      ItemData        =   "frmBatch.frx":000C
      Left            =   120
      List            =   "frmBatch.frx":000E
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   2580
      Width           =   2055
   End
   Begin VB.FileListBox filHeaders 
      Height          =   285
      Left            =   2280
      MultiSelect     =   2  'Extended
      Pattern         =   "*.RBH"
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   345
      Left            =   2280
      TabIndex        =   4
      Tag             =   "5003"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2280
      TabIndex        =   5
      Tag             =   "5004"
      Top             =   1515
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpenROM 
      Caption         =   "..."
      Height          =   285
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin VB.ListBox lstScripts 
      Height          =   1410
      ItemData        =   "frmBatch.frx":0010
      Left            =   120
      List            =   "frmBatch.frx":0012
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.FileListBox filScripts 
      Height          =   285
      Left            =   2280
      MultiSelect     =   2  'Extended
      Pattern         =   "*.RBC"
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Options"
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Tag             =   "5005"
      Top             =   3480
      Width           =   3375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   209
         TabIndex        =   16
         Top             =   240
         Width           =   3135
         Begin VB.CheckBox chkWithSTDItems 
            Caption         =   "Include STDItems"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Tag             =   "5007"
            Top             =   240
            Width           =   3015
         End
         Begin VB.CheckBox chkWithSTD 
            Caption         =   "Include STD"
            Height          =   255
            Left            =   0
            TabIndex        =   7
            Tag             =   "5006"
            Top             =   0
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CheckBox chkWithSTDPoke 
            Caption         =   "Include STDPoke"
            Height          =   255
            Left            =   0
            TabIndex        =   9
            Tag             =   "5008"
            Top             =   480
            Width           =   3015
         End
         Begin VB.CheckBox chkLog 
            Caption         =   "Show Log"
            Height          =   195
            Left            =   0
            TabIndex        =   11
            Tag             =   "5009"
            Top             =   990
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin VB.CheckBox chkWithSTDAttacks 
            Caption         =   "Include STDAttacks"
            Height          =   255
            Left            =   0
            TabIndex        =   10
            Tag             =   "5012"
            Top             =   720
            Width           =   3015
         End
      End
   End
   Begin VB.TextBox txtROM 
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblFiles 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RBC/RBH Files"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Tag             =   "5002"
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label lblROM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ROM"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Tag             =   "5001"
      Top             =   120
      Width           =   345
   End
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ListBox constants for SendMessage
Private Const LB_SETSEL = &H185
Private Const LB_GETSELCOUNT = &H190
Private Const LB_GETSELITEMS = &H191
Private Const LB_FINDSTRINGEXACT = &H1A2

Private Sub chkLog_Click()
    ' Invert the NoLog flag
    NoLog = Not NoLog
End Sub

Private Sub chkSelect_Click()
    
    ' Swap the caption value between Select All/None
    If chkSelect.Caption = LoadResString(5013) Then
        chkSelect.Caption = LoadResString(5014)
    Else
        chkSelect.Caption = LoadResString(5013)
    End If
    
    ' Select/deselect all items
    SendMessage lstScripts.hWnd, LB_SETSEL, chkSelect.Value, ByVal -1&
    SendMessage lstHeaders.hWnd, LB_SETSEL, chkSelect.Value, ByVal -1&
    
End Sub

Private Sub CompileCheck()
    
    ' If the ROM path isn't empty
    If LenB(txtROM.text) <> 0 Then
        
        If lstScripts.ListCount > 0 Then
            ' If some scripts were selected, enable Compile
            cmdCompile.Enabled = True
        Else
            ' Disable Compile otherwhise
            cmdCompile.Enabled = False
        End If
        
    Else
        ' Can't use Compile
        cmdCompile.Enabled = False
    End If

End Sub

Private Sub cmdBrowse_Click()
Dim sNewPath As String

    ' Set the paths
    sNewPath = BrowseFolder(Me.hWnd, LoadResString(5010))
    
    ' No path?
    If LenB(sNewPath) <> 0 Then
        
        ' Update the path is necessary
        If filScripts.Path <> sNewPath Then
        
            ' Set the new path
            filScripts.Path = sNewPath
            filHeaders.Path = sNewPath
        
        Else
            
            ' Otherwhise refresh
            filScripts.Refresh
            filHeaders.Refresh
            
        End If

        ' Fill the lists
        PopulateLists
    
    End If
    
End Sub

Private Sub cmdCompile_Click()
Dim i As Long
Dim lSelScripts() As Long
Dim lSelHeaders() As Long
Dim lScriptCount As Long
Dim lHeaderCount As Long
Dim sROM As String
    
    ' Get the number of selected items
    lScriptCount = SendMessage(lstScripts.hWnd, LB_GETSELCOUNT, 0&, ByVal 0&)
    lHeaderCount = SendMessage(lstHeaders.hWnd, LB_GETSELCOUNT, 0&, ByVal 0&)
    
    If lScriptCount <= 0 Then
        If lHeaderCount <= 0 Then
            ' If we reached this point, no file was selected
            ' No need to go ahead
            MsgBox LoadResString(5011), vbExclamation
            Exit Sub
        End If
    End If
    
    ' Set the mouse pointer to busy
    MousePointer = vbHourglass
    
    ' Update the NoLog flag
    NoLog = Not CBool(chkLog.Value)
    
    ' If the log is enabled, clean it
    If NoLog = False Then
        frmOutput.txtOutput.text = vbNullString
        frmOutput.lstDynamics.Clear
        frmOutput.lstOffsets.Clear
    End If
    
    ' Cleanup
    ClearData
      
    ' If the standard header was chosen, process it
    If chkWithSTD.Value = vbChecked Then
        HeaderProcess App.Path & "\std.rbh"
    End If
    
    ' If the item header was chosen, process it
    If chkWithSTDItems.Value = vbChecked Then
        HeaderProcess App.Path & "\stditems.rbh"
    End If
    
    ' If the Pokémon header was chosen, process it
    If chkWithSTDPoke.Value = vbChecked Then
        HeaderProcess App.Path & "\stdpoke.rbh"
    End If
    
    ' If the attack header was chosen, process it
    If chkWithSTDAttacks.Value = vbChecked Then
        HeaderProcess App.Path & "\stdattacks.rbh"
    End If
    
    ' Store the ROM path
    sROM = txtROM.text
    
    ' If some headers were selected
    If lHeaderCount > 0 Then
    
        ' Make the array large enough
        ReDim lSelHeaders(0 To lHeaderCount - 1) As Long
      
        ' Fill it with the selected indexes
        SendMessage lstHeaders.hWnd, LB_GETSELITEMS, lHeaderCount, lSelHeaders(0)
        
        ' Process the selected items
        For i = 0 To lHeaderCount - 1
            HeaderProcess filHeaders.Path & "\" & lstHeaders.List(lSelHeaders(i))
        Next i
    
    End If
    
    ' If some scripts were selected
    If lScriptCount > 0 Then
        
        ' Resize the array properly
        ReDim lSelScripts(0 To lScriptCount - 1) As Long
        
        ' Fill it with the selected indexes
        SendMessage lstScripts.hWnd, LB_GETSELITEMS, lScriptCount, lSelScripts(0)
        
        ' Process the selected items
        For i = 0 To lScriptCount - 1
            Process sROM, filScripts.Path & "\" & lstScripts.List(lSelScripts(i)), True
        Next i
    
    End If
    
    ' Restore the mouse
    MousePointer = vbDefault

End Sub

Private Sub cmdOpenROM_Click()
Dim oOpenDialog As clsCommonDialog
    
    ' Display an Open dialog and assign the result to the ROM TextBox
    Set oOpenDialog = New clsCommonDialog
    txtROM.text = oOpenDialog.ShowOpen(Me.hWnd, vbNullString, , "GameBoy Advance ROMs (*.gba)|*.gba|", FileMustExist Or PATHMUSTEXIST Or HideReadOnly)
    
    ' Free reference
    Set oOpenDialog = Nothing
    
End Sub

Private Sub PopulateLists()
Const LB_ERR = -1
Dim sToRemove(3) As String
Dim lRet() As Long
Dim i As Integer
    
    ' Clear ListBoxes
    lstScripts.Clear
    lstHeaders.Clear
    
    ' Lock redrawing
    LockUpdate lstScripts.hWnd
    
    ' Fill the script list using the FileListBox items
    For i = 0 To filScripts.ListCount - 1
        AddItem Me, lstScripts, filScripts.List(i)
    Next i
    
    ' Unlock redrawing
    UnlockUpdate lstScripts.hWnd
    
    ' If there's at least one item, select the first one
    If lstScripts.ListCount > 0 Then
        lstScripts.ListIndex = 0
    End If
    
    ' Lock redrawing
    LockUpdate lstHeaders.hWnd
    
    ' Fille the header list with the FileListBox files
    For i = 0 To filHeaders.ListCount - 1
        AddItem Me, lstHeaders, filHeaders.List(i)
    Next i
    
    ' Make sure there are no default headers
    sToRemove(0) = "std.rbh"
    sToRemove(1) = "stditems.rbh"
    sToRemove(2) = "stdpoke.rbh"
    sToRemove(3) = "stdattacks.rbh"
    
    ReDim lRet(UBound(sToRemove)) As Long
    
    For i = LBound(sToRemove) To UBound(sToRemove)
        
        lRet(i) = SendMessage(lstHeaders.hWnd, LB_FINDSTRINGEXACT, -1, ByVal sToRemove(i))
        
        If lRet(i) <> LB_ERR Then
            lstHeaders.RemoveItem (lRet(i))
        End If
        
    Next i
        
    ' Unlock redrawing
    UnlockUpdate lstHeaders.hWnd
    
    ' If there's at least one item, select the first one
    If lstHeaders.ListCount > 0 Then
        lstHeaders.ListIndex = 0
    End If
    
    ' Check if the Compile button should be enabled
    CompileCheck
    
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF5 Then
        
        ' Force refresh
        filScripts.Refresh
        filHeaders.Refresh
        
        ' Fill the lists
        PopulateLists
        
    End If

End Sub

Private Sub Form_Load()
Dim sTemp As String
   
    Localize Me
    
    ' Read the paths saved in the INI, if any
    sTemp = ReadIniString(App.Path & IniFile, "Batch", "FilePath", App.Path)
    
    ' If the path is not empty
    If LenB(sTemp) <> 0 Then
    
        ' Make sure the directory actually exists
        If LenB(Dir$(sTemp, vbDirectory)) Then
        
            filScripts.Path = sTemp
            filHeaders.Path = sTemp
            
            ' Fill the lists
            PopulateLists
            
        End If
    
    End If

    ' Read the ROM path stored into the INI, if any
    txtROM.text = ReadIniString(App.Path & IniFile, "Batch", "LastFile", vbNullString)
    
    ' Update the log CheckBox
    chkLog.Value = -CInt(Not NoLog)
    
    ' If the standard header exists
    If FileExists(App.Path & "\std.rbh") Then
        ' Read the value from the INI
        chkWithSTD.Value = ReadIniString(App.Path & IniFile, "Batch", "STD", 1)
    Else
        ' Otherwhise uncheck and disable
        chkWithSTD.Value = vbUnchecked
        chkWithSTD.Enabled = False
    End If
    
    ' If the item header exists
    If FileExists(App.Path & "\stditems.rbh") Then
        ' Read the value from the INI
        chkWithSTDItems.Value = ReadIniString(App.Path & IniFile, "Batch", "STDItems", 0)
    Else
        ' Otherwhise uncheck and disable
        chkWithSTDItems.Value = vbUnchecked
        chkWithSTDItems.Enabled = False
    End If
   
    ' If the Pokémon header exists
    If FileExists(App.Path & "\stdpoke.rbh") Then
        ' Read the value from the INI
        chkWithSTDPoke.Value = ReadIniString(App.Path & IniFile, "Batch", "STDPoke", 0)
    Else
        ' Otherwhise uncheck and disable
        chkWithSTDPoke.Value = vbUnchecked
        chkWithSTDPoke.Enabled = False
    End If
    
    ' If the attack header exists
    If FileExists(App.Path & "\stdattacks.rbh") Then
        ' Read the value from the INI
        chkWithSTDAttacks.Value = ReadIniString(App.Path & IniFile, "Batch", "STDAttacks", 0)
    Else
        ' Otherwhise uncheck and disable
        chkWithSTDAttacks.Value = vbUnchecked
        chkWithSTDAttacks.Enabled = False
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' Save all settings to the INI
    WriteStringToIni App.Path & IniFile, "Batch", "LastFile", txtROM.text
    WriteStringToIni App.Path & IniFile, "Batch", "FilePath", filScripts.Path
    WriteStringToIni App.Path & IniFile, "Batch", "STD", chkWithSTD.Value
    WriteStringToIni App.Path & IniFile, "Batch", "STDItems", chkWithSTDItems.Value
    WriteStringToIni App.Path & IniFile, "Batch", "STDPoke", chkWithSTDPoke.Value
    WriteStringToIni App.Path & IniFile, "Batch", "STDAttacks", chkWithSTDAttacks.Value
    
End Sub

Private Sub lstHeaders_Click()
    ' Manullay redraw the list to solve
    ' the refresh problem due to XP/Vista
    ' visual styles
    Redraw lstScripts.hWnd
End Sub

Private Sub lstHeaders_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ListBoxToolTip lstHeaders, Y
End Sub

Private Sub lstScripts_Click()
    ' Manullay redraw the list to solve
    ' the refresh problem due to XP/Vista
    ' visual styles
    Redraw lstScripts.hWnd
End Sub

Private Sub lstScripts_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ListBoxToolTip lstScripts, Y
End Sub

Private Sub txtROM_Change()
    ' Check if the Compile button should be enabled or not
    CompileCheck
End Sub
