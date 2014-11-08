VERSION 5.00
Begin VB.Form frmAssociate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Associate"
   ClientHeight    =   1095
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
   Icon            =   "frmAssociate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   73
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   377
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "4000"
   Begin VB.CheckBox chkIntegrate 
      Caption         =   "Associate GBA files with XSE"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "4013"
      Top             =   570
      Width           =   3855
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   345
      Left            =   4200
      TabIndex        =   2
      Tag             =   "4002"
      Top             =   150
      Width           =   1350
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4200
      TabIndex        =   3
      Tag             =   "4003"
      Top             =   592
      Width           =   1350
   End
   Begin VB.CheckBox chkAssoc 
      Caption         =   "Associate RBC/RBH/RBT files with XSE"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "4001"
      Top             =   270
      Width           =   3855
   End
End
Attribute VB_Name = "frmAssociate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sExePath As String
Private AlreadyAssociated As Boolean
Private AlreadyIntegrated As Boolean

Private Sub cmdApply_Click()
Dim sMessage As String
    
    If chkAssoc.Value = vbChecked Then
        ' Associate extensions, but only if they weren't already
        If AlreadyAssociated = False Then
            AssociateExt "rbc", sExePath, LoadResString(4005), 1
            AssociateExt "rbh", sExePath, LoadResString(4006), 2
            AssociateExt "rbt", sExePath, LoadResString(4007), 3
        End If
    Else
        ' Unassociate extensions, if they were associated previously
        If AlreadyAssociated = True Then
            UnAssociateExt "rbc", LoadResString(4005)
            UnAssociateExt "rbh", LoadResString(4006)
            UnAssociateExt "rbt", LoadResString(4007)
        End If
    End If
    
    If chkIntegrate.Value = vbChecked Then
        ' If XSE wasn't available into the .gba shell menu, add it
        If AlreadyIntegrated = False Then
            IntegrateShell "gba", sExePath, App.EXEName
        End If
    Else
        ' Remove XSE from the .gba shell menu, if necessary
        If AlreadyIntegrated = True Then
            UnintegrateShell "gba", App.EXEName
        End If
    End If
    
    ' Build the message string based on user's choices
    If chkAssoc.Value = vbChecked And chkIntegrate.Value = vbChecked Then
        sMessage = LoadResString(4011)
    ElseIf chkAssoc.Value = vbUnchecked And chkIntegrate.Value = vbUnchecked Then
        sMessage = LoadResString(4012)
    Else
        sMessage = "[RBC/RBH/RBT] " & LoadResString(4012 - chkAssoc.Value) & vbNewLine & _
        "[GBA] " & LoadResString(4012 - chkIntegrate.Value)
    End If
    
    ' Display the message and exit
    MsgBox sMessage, vbInformation
    Me.Hide
    
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function CheckRegKey(sSubKey As String, sFilePath As String) As Boolean
Dim sTemp As String
    
    ' Just in case some errors are raised
    On Error GoTo Hell
    
    ' Open the registry key
    RegOpenKey HKEY_CURRENT_USER, sUserClasses & sSubKey, phkResult
    
    ' Prepare a string buffer
    lpData = Space$(255)
    lpcbData = 255
    
    ' Finally get the value
    RegQueryValueEx phkResult, vbNullString, 0&, REG_SZ, lpData, lpcbData
    
    ' Set a temp string while removing quotes as well as the %1 part
    sTemp = Mid$(lpData, 2, lpcbData - 8)
    
    ' Close the registry key
    RegCloseKey (phkResult)
    
    ' If the temp string contains the file path, then it's valid
    CheckRegKey = InStrB(1, sTemp, sFilePath) <> 0
    
Hell:
End Function

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    Localize Me
            
    ' Define the full path
    sExePath = App.Path
    
    ' Make sure there's a slash
    If Right$(sExePath, 1) <> "\" Then
        sExePath = sExePath & "\"
    End If
    
    ' Add the exe name
    sExePath = sExePath & App.EXEName
    
    ' Make sure there's the extension
    If LCase$(Right$(sExePath, 4)) <> ".exe" Then
        sExePath = sExePath & ".exe"
    End If
    
    ' See if XSE was associated to script files and/or GBA ones
    lpSubKey = LoadResString(4005) & "\shell\open\command"
    chkAssoc.Value = -CInt(CheckRegKey(LoadResString(4005) & "\shell\open\command", sExePath))
    chkIntegrate.Value = -CInt(CheckRegKey(".gba" & "\shell\" & App.EXEName & "\command", sExePath))
    
    ' Update the internal flags accordingly
    AlreadyAssociated = chkAssoc.Value
    AlreadyIntegrated = chkIntegrate.Value
            
End Sub
