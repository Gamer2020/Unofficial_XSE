VERSION 5.00
Begin VB.Form frmGoto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Goto Line"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3015
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGoto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   79
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   201
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "3000"
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1560
      TabIndex        =   3
      Tag             =   "3003"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Tag             =   "3002"
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtLineNum 
      Height          =   285
      Left            =   1620
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line number"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Tag             =   "3001"
      Top             =   270
      Width           =   870
   End
End
Attribute VB_Name = "frmGoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lLimit As Long
Private Const ES_NUMBER As Long = &H2000

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    GotoLine (txtLineNum.text)
    Unload Me
End Sub

Public Sub GetLimit()

    ' Retrieve the actual amount of lines
    lLimit = SendMessage(Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_GETLINECOUNT, 0&, ByVal 0&)
    
    ' Set the MaxLength property accordingly
    txtLineNum.MaxLength = Len(CStr(lLimit))
    
    ' Get current line
    txtLineNum.text = SendMessage(Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_LINEFROMCHAR, -1, ByVal 0&) + 1
    txtLineNum.SelStart = 0
    txtLineNum.SelLength = Len(txtLineNum.text)
   
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    Localize Me
    
    ' Get the number limit
    GetLimit
    
    ' Make the textbox accept only numbers
    SetWindowLong txtLineNum.hWnd, GWL_STYLE, GetWindowLong(txtLineNum.hWnd, GWL_STYLE) Or ES_NUMBER
    
End Sub

Private Sub txtLineNum_Change()
    
    ' Check if it's empty or not
    If LenB(txtLineNum.text) <> 0 Then
        
        ' Not empty, enable the OK button
        cmdOK.Enabled = True
        
        ' If the value is higher than zero
        If CLng(txtLineNum.text) > 0 Then
            
            ' Make sure it's not past the limit
            If CLng(txtLineNum.text) > lLimit Then
                txtLineNum.text = lLimit
                txtLineNum.SelStart = Len(CStr(lLimit))
            End If
            
        Else
            ' Value is zero, so fix it
            txtLineNum.text = 1
            txtLineNum.SelStart = Len(txtLineNum.text)
        End If
        
    Else
        ' No value, no party
        cmdOK.Enabled = False
    End If
    
End Sub

Private Sub txtLineNum_KeyPress(KeyCode As Integer)
    
    ' Pressing Enter will trigger the OK button
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        cmdOK_Click
    End If
    
End Sub
