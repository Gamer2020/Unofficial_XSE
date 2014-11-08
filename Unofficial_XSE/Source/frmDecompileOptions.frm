VERSION 5.00
Begin VB.Form frmDecompileOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Decompile Options"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDecompileOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   181
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   328
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "15000"
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3480
      TabIndex        =   10
      Tag             =   "4003"
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Frame fraRefactoring 
      Caption         =   "Refactoring"
      Height          =   735
      Left            =   2520
      TabIndex        =   15
      Tag             =   "15007"
      Top             =   1200
      Width           =   2295
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   137
         TabIndex        =   16
         Top             =   240
         Width           =   2055
         Begin VB.TextBox txtPrefix 
            Enabled         =   0   'False
            Height          =   285
            Left            =   870
            TabIndex        =   18
            Text            =   "0x"
            Top             =   60
            Width           =   255
         End
         Begin VB.TextBox txtDynamic 
            Height          =   285
            Left            =   1110
            MaxLength       =   7
            TabIndex        =   8
            Top             =   60
            Width           =   870
         End
         Begin VB.Label lblDynamic 
            AutoSize        =   -1  'True
            Caption         =   "#dynamic"
            Enabled         =   0   'False
            Height          =   195
            Left            =   60
            TabIndex        =   17
            Top             =   90
            Width           =   705
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   345
      Left            =   2040
      TabIndex        =   9
      Tag             =   "4002"
      Top             =   2280
      Width           =   1350
   End
   Begin VB.Frame Frame2 
      Caption         =   "Decompile Mode"
      Height          =   2055
      Left            =   120
      TabIndex        =   13
      Tag             =   "15001"
      Top             =   120
      Width           =   2295
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   14
         Top             =   240
         Width           =   2115
         Begin VB.CheckBox chkComments 
            Caption         =   "Comments"
            Height          =   195
            Left            =   120
            TabIndex        =   3
            Tag             =   "15006"
            Top             =   1080
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chkRefactor 
            Caption         =   "Refactoring"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Tag             =   "15007"
            Top             =   1380
            Width           =   1575
         End
         Begin VB.OptionButton optMode 
            Caption         =   "Strict"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Tag             =   "15004"
            Top             =   690
            Width           =   1845
         End
         Begin VB.OptionButton optMode 
            Caption         =   "Normal"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   1
            Tag             =   "15003"
            Top             =   390
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VB.OptionButton optMode 
            Caption         =   "Enhanced"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   0
            Tag             =   "15002"
            Top             =   90
            Width           =   1845
         End
      End
   End
   Begin VB.Frame fraComments 
      Caption         =   "Comments to use"
      Height          =   735
      Left            =   2520
      TabIndex        =   11
      Tag             =   "15005"
      Top             =   360
      Width           =   2295
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   320
         ScaleHeight     =   25
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   113
         TabIndex        =   12
         Top             =   240
         Width           =   1695
         Begin VB.OptionButton optComment 
            Caption         =   " //"
            Height          =   255
            Index           =   2
            Left            =   1170
            TabIndex        =   7
            Top             =   90
            Width           =   495
         End
         Begin VB.OptionButton optComment 
            Caption         =   " ;"
            Height          =   255
            Index           =   1
            Left            =   660
            TabIndex        =   6
            Top             =   90
            Width           =   495
         End
         Begin VB.OptionButton optComment 
            Caption         =   " '"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   90
            Value           =   -1  'True
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frmDecompileOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private hexDynamic As clsHexBox

Private Function GetSelectedOpt(optButtons As Variant) As Integer
Dim i As Integer
    
    ' Loop through all the options
    For i = 0 To optButtons.Count - 1
        ' If the option is true
        If optButtons(i).Value = True Then
            ' Then we got the selected one
            Exit For
        End If
    Next i
    
    GetSelectedOpt = i

End Function

Private Sub chkComments_Click()
    
    ' Enable/disable the Comments frame and its options
    fraComments.Enabled = CBool(chkComments.Value)
    optComment(0).Enabled = CBool(chkComments.Value)
    optComment(1).Enabled = CBool(chkComments.Value)
    optComment(2).Enabled = CBool(chkComments.Value)
    
End Sub

Private Sub chkRefactor_Click()
    
    ' Enabled/disable the Refactoring frame and its content
    fraRefactoring.Enabled = CBool(chkRefactor.Value)
    lblDynamic.Enabled = CBool(chkRefactor.Value)
    txtDynamic.Enabled = CBool(chkRefactor.Value)
    
End Sub

Private Sub cmdApply_Click()
    
    ' Update all the needed variables
    iDecompileMode = GetSelectedOpt(optMode)
    iComments = chkComments.Value
    iRefactoring = chkRefactor.Value
    sCommentChar = optComment(GetSelectedOpt(optComment)).Caption
    sRefactorDynamic = txtDynamic.text
    
    ' Save settings to INI
    WriteStringToIni App.Path & IniFile, "Options", "DecompileMode", iDecompileMode
    WriteStringToIni App.Path & IniFile, "Options", "Comments", iComments
    WriteStringToIni App.Path & IniFile, "Options", "Refactoring", iRefactoring
    WriteStringToIni App.Path & IniFile, "Options", "CommentChar", sCommentChar
    WriteStringToIni App.Path & IniFile, "Options", "RefactorDynamic", sRefactorDynamic
    
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    Localize Me
    
    ' Subclass the TextBox
    Set hexDynamic = New clsHexBox
    Set hexDynamic.TextBox = txtDynamic
    
    ' Read settings from the INI
    optMode(ReadIniString(App.Path & IniFile, "Options", "DecompileMode", 1)).Value = True
    chkComments.Value = ReadIniString(App.Path & IniFile, "Options", "Comments", 1)
    chkRefactor.Value = ReadIniString(App.Path & IniFile, "Options", "Refactoring", 0)
    txtDynamic.text = ReadIniString(App.Path & IniFile, "Options", "RefactorDynamic")
    txtDynamic.Enabled = CBool(chkRefactor.Value)
    
    Select Case ReadIniString(App.Path & IniFile, "Options", "CommentChar", "'")
        Case "'"
            optComment(0).Value = True
        Case ";"
            optComment(1).Value = True
        Case "//"
            optComment(2).Value = True
        Case Else
            optComment(1).Value = True
    End Select
    
End Sub

Private Sub txtDynamic_Change()
    
    ' If the file size isn't higher than 16 MB
    If FileLength(Document(frmMain.Tabs.SelectedTab).LoadedFile) <= &H1000000 Then
        ' No offset can go past 6 chars
        txtDynamic.MaxLength = 6
        txtDynamic.text = Left$(txtDynamic.text, 6)
    Else
        ' Else we need 7 chars
        txtDynamic.MaxLength = 7
    End If
    
    ' Make sure it has at least 6 chars
    If Len(txtDynamic.text) >= 6 Then
        
        ' Check if it's a valid hex number
        If IsHex(txtDynamic.text) Then
            
            ' If the TextBox is not enabled
            If txtDynamic.Enabled = False Then
                ' Reset the text
                txtDynamic.text = vbNullString
            End If
            
        End If

    End If
    
End Sub
