VERSION 5.00
Begin VB.Form frmOutput 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Compiler Output"
   ClientHeight    =   5880
   ClientLeft      =   7320
   ClientTop       =   3270
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOutput.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "11000"
   Begin VB.Frame fraDynamic 
      Caption         =   "Dynamic Offsets"
      Height          =   1620
      Left            =   120
      TabIndex        =   4
      Tag             =   "11003"
      Top             =   4140
      Width           =   4335
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   81
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   275
         TabIndex        =   5
         Top             =   240
         Width           =   4120
         Begin VB.ListBox lstOffsets 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            ItemData        =   "frmOutput.frx":000C
            Left            =   1500
            List            =   "frmOutput.frx":000E
            TabIndex        =   8
            Top             =   60
            Width           =   1095
         End
         Begin VB.ListBox lstDynamics 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            ItemData        =   "frmOutput.frx":0010
            Left            =   60
            List            =   "frmOutput.frx":0012
            TabIndex        =   7
            Top             =   60
            Width           =   1335
         End
         Begin VB.CommandButton cmdCopy 
            Caption         =   "Copy"
            Enabled         =   0   'False
            Height          =   345
            Left            =   2730
            TabIndex        =   6
            Tag             =   "11004"
            Top             =   60
            Width           =   1275
         End
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   345
      Left            =   330
      TabIndex        =   2
      Tag             =   "57"
      Top             =   3720
      Width           =   1275
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   345
      Left            =   1650
      TabIndex        =   3
      Tag             =   "11001"
      Top             =   3720
      Width           =   1275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save..."
      Height          =   345
      Left            =   2970
      TabIndex        =   0
      Tag             =   "11002"
      Top             =   3720
      Width           =   1275
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsLinking As Boolean

Private Sub cmdClear_Click()
    txtOutput.text = vbNullString
End Sub

Private Sub cmdClose_Click()
    If GetActiveWindow = Me.hWnd Then
        Me.Hide
    Else
        Unload Me
    End If
End Sub

Private Sub CopyCheck()
    If IsHex(lstOffsets.List(lstOffsets.ListIndex)) Then
        cmdCopy.Enabled = True
    Else
        cmdCopy.Enabled = False
    End If
End Sub

Private Sub cmdCopy_Click()
    SafeClipboardSet lstOffsets.List(lstOffsets.ListIndex)
End Sub

Private Sub cmdSave_Click()
Dim sResult As String
Dim iFileNum As Integer
Dim oOpenDialog As clsCommonDialog
    
    Set oOpenDialog = New clsCommonDialog
    sResult = oOpenDialog.ShowSave(Me.hWnd, vbNullString, , , "Compiler Log (*.log)|*.log|", OVERWRITEPROMPT Or PATHMUSTEXIST)
    Set oOpenDialog = Nothing
    
    If LenB(sResult) <> 0 Then
    
        iFileNum = FreeFile
        
        Open sResult For Output As #iFileNum
            Print #iFileNum, txtOutput.text;
        Close #iFileNum
        
    End If
    
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Localize Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If GetActiveWindow = Me.hWnd Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub lstDynamics_DblClick()
    cmdCopy_Click
End Sub

Private Sub lstDynamics_Click()
    
    If IsLinking = False Then
        
        IsLinking = True
        
        lstOffsets.TopIndex = lstDynamics.TopIndex
        lstOffsets.ListIndex = lstDynamics.ListIndex
        CopyCheck
        
        IsLinking = False
        
    End If
    
End Sub

Private Sub lstDynamics_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ListBoxToolTip lstDynamics, Y
End Sub

Private Sub lstDynamics_Scroll()
    
    If IsLinking = False Then
        
        IsLinking = True
        
        lstOffsets.TopIndex = lstDynamics.TopIndex
        lstOffsets.ListIndex = lstDynamics.ListIndex
        CopyCheck
        
        IsLinking = False
        
    End If
    
End Sub

Private Sub lstOffsets_Click()
    
    If IsLinking = False Then
        
        IsLinking = True
        
        lstDynamics.TopIndex = lstOffsets.TopIndex
        lstDynamics.ListIndex = lstOffsets.ListIndex
        CopyCheck
        
        IsLinking = False
        
    End If
    
End Sub

Private Sub lstOffsets_DblClick()
    cmdCopy_Click
End Sub

Private Sub lstOffsets_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ListBoxToolTip lstOffsets, Y
End Sub

Private Sub lstOffsets_Scroll()
    
    If IsLinking = False Then
        
        IsLinking = True
        
        lstDynamics.TopIndex = lstOffsets.TopIndex
        lstDynamics.ListIndex = lstOffsets.ListIndex
        CopyCheck
        
        IsLinking = False
        
    End If
    
End Sub

Private Sub txtOutput_Change()
    If LenB(txtOutput.text) <> 0 Then
        cmdClear.Enabled = True
        cmdSave.Enabled = True
    Else
        cmdClear.Enabled = False
        cmdSave.Enabled = False
    End If
End Sub
