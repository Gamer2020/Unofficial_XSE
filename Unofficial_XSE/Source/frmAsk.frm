VERSION 5.00
Begin VB.Form frmAsk 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAsk.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   93
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   496
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "Cancel"
      Height          =   345
      Index           =   4
      Left            =   5925
      TabIndex        =   4
      Tag             =   "1016"
      Top             =   900
      Width           =   1350
   End
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "No to All"
      Height          =   345
      Index           =   3
      Left            =   4485
      TabIndex        =   3
      Tag             =   "1015"
      Top             =   900
      Width           =   1350
   End
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "No"
      Height          =   345
      Index           =   2
      Left            =   3045
      TabIndex        =   2
      Tag             =   "1014"
      Top             =   900
      Width           =   1350
   End
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "Yes to All"
      Height          =   345
      Index           =   1
      Left            =   1605
      TabIndex        =   1
      Tag             =   "1013"
      Top             =   900
      Width           =   1350
   End
   Begin VB.CommandButton cmdPrompt 
      Caption         =   "Yes"
      Height          =   345
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Tag             =   "1012"
      Top             =   900
      Width           =   1350
   End
   Begin VB.Label lblPrompt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Do you want to save the changes made to the file?"
      Height          =   195
      Left            =   1920
      TabIndex        =   5
      Tag             =   "1001"
      Top             =   300
      Width           =   3705
   End
End
Attribute VB_Name = "frmAsk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const IDI_EXCLAMATION   As Long = &H7F03&
Private Const MB_ICONEXCLAMATION As Long = &H30&

Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, lpIconName As Any) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Private Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long

Private Sub cmdPrompt_Click(Index As Integer)
    
    ' If the last button wasn't clicked
    If Index < cmdPrompt.Count - 1 Then
        ' Set the answer as the index value +1
        frmMain.Answer = Index + 1
    Else
        ' Else Cancel was pressed, so the answer is zero
        frmMain.Answer = 0
    End If
    
    ' In any case, unload
    Unload Me
    
End Sub

Private Sub Form_Activate()
    ' Play the exclamation sound
    MessageBeep MB_ICONEXCLAMATION
End Sub

Private Sub Form_Load()
Dim hIcon As Long
    
    Localize Me
    Me.Caption = App.Title
    
    ' Load the proper icon resource from the system
    hIcon = LoadIcon(0&, ByVal IDI_EXCLAMATION)
    
    ' Draw the icon and free the memory
    DrawIcon Me.hDC, 12, 12, hIcon
    DestroyIcon hIcon

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' If the user didn't press any button at this point
    If frmMain.Answer = -1 Then
        ' Set the answer to Cancel
        frmMain.Answer = 0
    End If
    
End Sub
