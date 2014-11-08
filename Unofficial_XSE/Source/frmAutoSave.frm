VERSION 5.00
Begin VB.Form frmAutoSave 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Auto Save"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutoSave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   93
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "17000"
   Begin VB.HScrollBar hsbInterval 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   300
      Min             =   10
      SmallChange     =   10
      TabIndex        =   1
      Top             =   480
      Value           =   60
      Width           =   2415
   End
   Begin VB.TextBox txtInterval 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "60"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   1920
      TabIndex        =   3
      Tag             =   "3002"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3240
      TabIndex        =   4
      Tag             =   "3003"
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox chkSave 
      Caption         =   "Save RBC/RBH/RBT files automatically every:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Tag             =   "17001"
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblSeconds 
      AutoSize        =   -1  'True
      Caption         =   " seconds."
      Enabled         =   0   'False
      Height          =   195
      Left            =   3240
      TabIndex        =   5
      Tag             =   "13025"
      Top             =   510
      Width           =   690
   End
End
Attribute VB_Name = "frmAutoSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ES_NUMBER As Long = &H2000
Private IsScrolling As Boolean

Private Sub chkSave_Click()
    hsbInterval.Enabled = CBool(chkSave.Value)
    txtInterval.Enabled = CBool(chkSave.Value)
    lblSeconds.Enabled = CBool(chkSave.Value)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If chkSave.Value = vbChecked Then
        frmMain.objTmr.StartTimer 1000& * hsbInterval.Value
    Else
        frmMain.objTmr.StopTimer
    End If
    
    frmMain.mnuAutoSave.Checked = CBool(chkSave.Value)
    
    WriteStringToIni App.Path & IniFile, "Options", "AutoSave", chkSave.Value
    WriteStringToIni App.Path & IniFile, "Options", "SaveInterval", hsbInterval.Value
    
    Unload Me
    
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    Localize Me
    
    chkSave.Value = ReadIniString(App.Path & IniFile, "Options", "AutoSave", 0)
    hsbInterval.Value = ReadIniString(App.Path & IniFile, "Options", "SaveInterval", 60)
    
    ' Make the textbox accept only numbers
    SetWindowLong txtInterval.hWnd, GWL_STYLE, GetWindowLong(txtInterval.hWnd, GWL_STYLE) Or ES_NUMBER
    
End Sub

Private Sub hsbInterval_Change()
    hsbInterval_Scroll
End Sub

Private Sub hsbInterval_Scroll()
    
    If IsScrolling = False Then
        IsScrolling = True
        txtInterval.text = hsbInterval.Value
        txtInterval.SelStart = Len(txtInterval.text)
        IsScrolling = False
    End If
    
End Sub

Private Sub txtInterval_Change()

    On Error GoTo Hell

    If IsScrolling = False Then
        
        IsScrolling = True
        
        If LenB(txtInterval.text) <> 0 Then
            
            If CInt(txtInterval.text) < hsbInterval.Min Then
                txtInterval.ForeColor = vbRed
            ElseIf CInt(txtInterval.text) > hsbInterval.Max Then
                txtInterval.ForeColor = vbRed
            Else
                txtInterval.ForeColor = vbWindowText
            End If
            
            hsbInterval.Value = CInt(txtInterval.text)
            
        End If
        
        IsScrolling = False
        
    End If
    
    Exit Sub

Hell:
IsScrolling = False
End Sub

Private Sub txtInterval_LostFocus()
    
    If LenB(txtInterval.text) Then
    
        If CInt(txtInterval.text) < hsbInterval.Min Then
            txtInterval.text = hsbInterval.Min
            txtInterval.SelStart = Len(txtInterval.text)
        ElseIf CInt(txtInterval.text) > hsbInterval.Max Then
            txtInterval.text = hsbInterval.Max
            txtInterval.SelStart = Len(txtInterval.text)
        End If
    
    End If
End Sub
