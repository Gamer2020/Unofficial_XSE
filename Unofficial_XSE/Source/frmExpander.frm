VERSION 5.00
Begin VB.Form frmExpander 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ROM Resizer"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExpander.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   217
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "16000"
   Begin VB.Frame Frame2 
      Caption         =   "Shrink to"
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Tag             =   "16006"
      Top             =   1320
      Width           =   3015
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   185
         TabIndex        =   9
         Top             =   240
         Width           =   2775
         Begin VB.OptionButton optSize 
            Caption         =   "8 MB"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   1
            Tag             =   "16002"
            Top             =   90
            Width           =   735
         End
         Begin VB.OptionButton optSize 
            Caption         =   "16 MB"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   2
            Tag             =   "16007"
            Top             =   90
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.CommandButton cmdShrink 
            Caption         =   "Shrink"
            Height          =   345
            Left            =   1500
            TabIndex        =   3
            Tag             =   "16008"
            Top             =   450
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Expand"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Tag             =   "16001"
      Top             =   60
      Width           =   3015
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   120
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   185
         TabIndex        =   7
         Top             =   240
         Width           =   2775
         Begin VB.CommandButton cmdExpand 
            Caption         =   "Expand"
            Height          =   345
            Left            =   1500
            TabIndex        =   0
            Tag             =   "16005"
            Top             =   450
            Width           =   1215
         End
         Begin VB.ComboBox cboFillByte 
            Height          =   315
            ItemData        =   "frmExpander.frx":000C
            Left            =   120
            List            =   "frmExpander.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   120
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fill Byte"
            Height          =   195
            Left            =   840
            TabIndex        =   8
            Tag             =   "16004"
            Top             =   165
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "frmExpander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' File Constants
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_BEGIN = 0

' File handling APIs
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Private Sub TruncateFile(sFileName As String, ByVal lSize As Long)
Dim hFile As Long
    
    ' Get a handle for the file
    hFile = CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
    
    ' Seek the location to then new size
    SetFilePointer hFile, lSize, 0, FILE_BEGIN
    
    ' Set the end of file there and close the file
    SetEndOfFile hFile
    CloseHandle hFile

End Sub

Private Sub cmdExpand_Click()
Dim iFileNum As Integer
Dim bFiller() As Byte
Dim lFileLen As Long
Dim lNewSize As Long
    
    On Error GoTo Finish
    
    ' Make sure the fill byte is selected
    If cboFillByte.ListIndex < 0 Then cboFillByte.ListIndex = 1
    
    ' Get a fresh new handle
    iFileNum = FreeFile
    
    ' Open the ROM
    Open Document(frmMain.Tabs.SelectedTab).LoadedFile For Binary As #iFileNum
        
        ' 32 MB
        lNewSize = &H2000000
        
        ' Check the size of the ROM
        lFileLen = LOF(iFileNum)
        
        ' If it can be expanded
        If lFileLen < lNewSize Then
            
            ' Set the mouse to busy
            Screen.MousePointer = vbHourglass
            
            ' Redim the array to met the new size
            If cboFillByte.ListIndex = 1 Then
                ReDim bFiller((lNewSize - 1) - lFileLen)
            End If
            
        Else
            ' No need to expand
            MsgBox LoadResString(16009), vbExclamation
            Exit Sub
        End If
        
        ' if FF is selected, fill the array accordingly
        If cboFillByte.ListIndex = 1 Then
            RtlFillMemory bFiller(0), UBound(bFiller) + 1, &HFF
        End If
                    
        ' Write the array to the ROM, if needed
        If cboFillByte.ListIndex = 1 Then
            Put #iFileNum, lFileLen + 1, bFiller
        Else
            Put #iFileNum, lNewSize, CByte(&H0)
        End If
        
        ' Done!
        MsgBox LoadResString(16010) & " (32 MB).", vbInformation
        Me.Hide
        
Finish:

    Close #iFileNum

    ' Reset mouse and exit
    Screen.MousePointer = vbDefault
    
    If Err.Number <> 0 Then
        MsgBox LoadResString(10026) & Err.Number & ": " & Err.Description & ".", vbExclamation
    End If
    
    Unload Me

End Sub

Private Sub cmdShrink_Click()
Dim iFileNum As Integer
Dim sFileName As String
Dim lFileLen As Long
Dim lOldSize As Long
    
    On Error GoTo Finish
    
    ' Start getting a new handle
    iFileNum = FreeFile
    
    ' Get the file name
    sFileName = Document(frmMain.Tabs.SelectedTab).LoadedFile
    
    ' Access the ROM
    Open sFileName For Binary As #iFileNum
        
        If optSize(1).Value = True Then
            ' 16 MB
            lOldSize = &H1000000
        ElseIf optSize(0).Value = True Then
             ' 8 MB
            lOldSize = &H800000
        End If
        
        ' Check the actual size of the ROM
        lFileLen = LOF(iFileNum)
        
        If lFileLen > lOldSize Then
            
            ' Safety prompt
            If MsgBox(LoadResString(16011), vbExclamation + vbYesNo) = vbNo Then
                Exit Sub
            End If
            
            ' Set the mouse to busy
            Screen.MousePointer = vbHourglass
            
            TruncateFile sFileName, lOldSize
                
        Else
            MsgBox LoadResString(16012), vbExclamation
            Exit Sub
        End If

        ' Finished
        MsgBox LoadResString(16013) & " (" & lOldSize \ 1048576 & " MB).", vbInformation
        Me.Hide

Finish:

    Close #iFileNum

    ' Reset mouse and exit
    Screen.MousePointer = vbDefault
    
    If Err.Number <> 0 Then
        MsgBox LoadResString(10026) & Err.Number & ": " & Err.Description & ".", vbExclamation
    End If
    
    Unload Me
    
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    Localize Me
    
    ' Initialize ListBox
    cboFillByte.ListIndex = 1
    
End Sub

Public Sub ToggleEnable(sFile As String)
Dim i As Long
    
    ' If the file is a ROM
    If LenB(sFile) <> 0 And GetExt(sFile) = "gba" Then
        For i = 0 To Me.Controls.Count - 1
            ' Make sure the controls are enabled
            Me.Controls(i).Enabled = True
        Next i
    Else
        For i = 0 To Me.Controls.Count - 1
            ' Make them disabled
            Me.Controls(i).Enabled = False
        Next i
    End If
    
End Sub
