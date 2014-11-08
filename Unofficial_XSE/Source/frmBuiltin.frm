VERSION 5.00
Begin VB.Form frmBuiltin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Script Templates"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   315
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
   Icon            =   "frmBuiltin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   246
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   241
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "6000"
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Tag             =   "6004"
      Top             =   2250
      Width           =   3375
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   209
         TabIndex        =   5
         Top             =   240
         Width           =   3135
         Begin VB.TextBox txtPreview 
            ForeColor       =   &H80000011&
            Height          =   840
            Left            =   60
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   60
            Width           =   3000
         End
      End
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   345
      Left            =   2280
      TabIndex        =   2
      Tag             =   "6002"
      Top             =   570
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   345
      Left            =   2280
      TabIndex        =   1
      Tag             =   "6001"
      Top             =   120
      Width           =   1215
   End
   Begin VB.FileListBox filFiles 
      Height          =   285
      Left            =   2280
      MultiSelect     =   2  'Extended
      Pattern         =   "*.BIC"
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox lstFiles 
      Height          =   2010
      ItemData        =   "frmBuiltin.frx":000C
      Left            =   120
      List            =   "frmBuiltin.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmBuiltin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const LB_ADDSTRING = &H180

Private Sub InsertCheck()
    
    ' If there are some items
    If lstFiles.ListCount > 0 Then
        ' Enable Insert
        cmdInsert.Enabled = True
    Else
        ' Disable it otherwhise
        cmdInsert.Enabled = False
    End If
    
End Sub

Private Sub cmdBrowse_Click()
Dim i As Integer
    
    ' Get the path
    filFiles.Path = BrowseFolder(Me.hWnd, LoadResString(6003))
    
    ' Clean both preview and the ListBox
    txtPreview.text = vbNullString
    lstFiles.Clear
    
    For i = 0 To filFiles.ListCount - 1
        ' Add each file to the list
        SendMessageStr lstFiles.hWnd, LB_ADDSTRING, 0, ByVal filFiles.List(i)
    Next i
    
    ' Check if Insert should be enabled or not
    InsertCheck

End Sub

Private Sub cmdInsert_Click()
Dim iFileNum As Integer
Dim sBuffer As String
    
    ' If the ListIndex is properly set
    If lstFiles.ListIndex >= 0 Then
        
        ' Get a new handle
        iFileNum = FreeFile
        
        ' Open the selected file
        Open filFiles.Path & "\" & lstFiles.List(lstFiles.ListIndex) For Binary As #iFileNum
            
            ' Initialize the buffer
            sBuffer = SysAllocStringLen(vbNullString, LOF(iFileNum))
            
            ' Put the file content into the buffer
            Get #iFileNum, 1, sBuffer
            
        Close #iFileNum
        
        ' Replace the current selection with the buffer
        SendMessageStr Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_REPLACESEL, True, sBuffer
        
    End If
    
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
    
    LocalizeStrings Me
    
    ' Retrieve the path form the INI
    filFiles.Path = ReadIniString(App.Path & IniFile, "Built-in", "LastPath", App.Path)
    
    LockUpdate lstFiles.hWnd
    
    ' Fill the list
    For i = 0 To filFiles.ListCount - 1
        SendMessageStr lstFiles.hWnd, LB_ADDSTRING, 0, ByVal filFiles.List(i)
    Next i
    
    UnlockUpdate lstFiles.hWnd
    lstFiles.ToolTipText = lstFiles.List(lstFiles.ListIndex)
    
    ' Check if Insert should be enabled or not
    InsertCheck
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' Save settings to INI
    WriteStringToIni App.Path & IniFile, "Built-in", "LastPath", filFiles.Path
    
End Sub

Private Sub lstFiles_Click()
Dim iFileNum As Integer
Dim sPreview As String * 256
    
    ' If the ListIndex is properly set
    If lstFiles.ListIndex >= 0 Then
    
        ' Get a new file handle
        iFileNum = FreeFile
        
        ' Open the selected file
        Open filFiles.Path & "\" & lstFiles.List(lstFiles.ListIndex) For Binary As #iFileNum
            ' Get a small preview
            Get #iFileNum, 1, sPreview
        Close #iFileNum
        
        ' Display the preview content
        txtPreview.text = Left$(sPreview, Len(sPreview) - 1)
        
    End If
    
    lstFiles.ToolTipText = lstFiles.List(lstFiles.ListIndex)
    
End Sub

Private Sub lstFiles_DblClick()
    cmdInsert_Click
End Sub

Private Sub lstFiles_Scroll()
    lstFiles.ToolTipText = lstFiles.List(lstFiles.ListIndex)
End Sub
