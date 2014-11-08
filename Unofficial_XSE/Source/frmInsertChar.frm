VERSION 5.00
Begin VB.Form frmInsertChar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Special Characters"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInsertChar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   111
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   137
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "8000"
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   960
      TabIndex        =   2
      Tag             =   "8002"
      Top             =   540
      Width           =   975
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   345
      Left            =   960
      TabIndex        =   1
      Tag             =   "8001"
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox lstChars 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      ItemData        =   "frmInsertChar.frx":000C
      Left            =   120
      List            =   "frmInsertChar.frx":009A
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmInsertChar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    SendMessageStr frmTextAdjuster.txtToAdjust.hWnd, EM_REPLACESEL, 1&, Left$(lstChars.List(lstChars.ListIndex), 1)
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Localize Me
End Sub

Private Sub lstChars_DblClick()
    cmdInsert_Click
End Sub
