VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2100
   ClientLeft      =   6570
   ClientTop       =   2565
   ClientWidth     =   6180
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   140
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   412
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "2000"
   Begin VB.CommandButton cmdWeb 
      Caption         =   "Web"
      Height          =   345
      Left            =   3720
      TabIndex        =   1
      Top             =   1620
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   4920
      TabIndex        =   0
      Top             =   1620
      Width           =   1140
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      Caption         =   "  Copyright  "
      ForeColor       =   &H80000011&
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   1365
      Width           =   885
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descr"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Tag             =   "2001"
      Top             =   960
      Width           =   405
   End
   Begin VB.Line linShadow 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   415
      Y1              =   49
      Y2              =   49
   End
   Begin VB.Line linCopyright 
      BorderColor     =   &H80000010&
      X1              =   6
      X2              =   406
      Y1              =   98
      Y2              =   98
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   150
      Left            =   5640
      TabIndex        =   2
      Tag             =   "0"
      Top             =   480
      Width           =   390
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblAppName 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "AppName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   960
      TabIndex        =   4
      Top             =   210
      Width           =   1065
   End
   Begin VB.Shape shpHeaderBackground 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   6225
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Copyright © 2009 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Private Const sMyName As String = "frmAbout"
Private Const ID_ABOUTPICTURE As Long = 102&

Private Declare Function ShellExecuteW Lib "shell32" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long

Private Sub cmdOK_Click()
Const sThis As String = "cmdOK_Click"
    
    On Error GoTo LocalHandler
    
    ' Unload the form
    Unload Me
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
    
End Sub

Private Sub cmdWeb_Click()
    
    ' Load the web page in the browser
    ShellExecuteW Me.hWnd, StrPtr("open"), StrPtr("http://www.andreasartori.net/hackmew"), 0&, 0&, vbNormalFocus
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Const sThis As String = "Form_KeyPress"
    
    On Error GoTo LocalHandler
    
    ' Mimic Windows' usual behaviour
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select

End Sub

Private Sub Form_Load()
Const sThis As String = "Form_Load"
Const sDot As String = "."
Const sSpace As String = " "
    
    On Error GoTo LocalHandler
    
    ' Localize the form
    Localize Me
    
    ' Set the name, version and copyright labels
    lblAppName.Caption = App.ProductName
    lblVersion.Caption = "v" & App.Major & sDot & App.Minor & sDot & App.Revision
    lblCopyright.Caption = sSpace & App.LegalCopyright & sSpace
    
    ' Load the About picture
    Set imgIcon.Picture = LoadResPicture(ID_ABOUTPICTURE, vbResBitmap)
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    ' Free the memory associated with the form
    Set frmAbout = Nothing
    
End Sub
