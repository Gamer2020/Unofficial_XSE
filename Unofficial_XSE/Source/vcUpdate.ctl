VERSION 5.00
Begin VB.UserControl vcUpdate 
   ClientHeight    =   3690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5985
   DataBindingBehavior=   1  'vbSimpleBound
   DataSourceBehavior=   1  'vbDataSource
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Picture         =   "vcUpdate.ctx":0000
   ScaleHeight     =   246
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ToolboxBitmap   =   "vcUpdate.ctx":1350
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3330
      TabIndex        =   2
      Tag             =   "10010"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >>"
      Height          =   345
      Left            =   4650
      TabIndex        =   1
      Tag             =   "10011"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Step 1 of 3"
      Height          =   2055
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Tag             =   "10001"
      Top             =   1080
      Width           =   5775
      Begin VB.Label lblStep1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"vcUpdate.ctx":1662
         Height          =   975
         Left            =   240
         TabIndex        =   3
         Tag             =   "10002"
         Top             =   360
         Width           =   5295
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Step 3 of 3"
      Height          =   2055
      Index           =   2
      Left            =   90
      TabIndex        =   10
      Tag             =   "10007"
      Top             =   1080
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label lblUpdate2 
         BackStyle       =   0  'Transparent
         Caption         =   "Destination: $f"
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Tag             =   "10009"
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label lblUpdate 
         BackStyle       =   0  'Transparent
         Caption         =   "Downloading update for $n..."
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Tag             =   "10008"
         Top             =   360
         Width           =   5415
      End
      Begin VB.Label lblProgressUpdate 
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   5415
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Step 2 of 3"
      Height          =   2055
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Tag             =   "10003"
      Top             =   1080
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label lblNo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "No update available. Click Cancel to exit."
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Tag             =   "10005"
         Top             =   1320
         Visible         =   0   'False
         Width           =   5295
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label lblProgress 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " "
         ForeColor       =   &H80000011&
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   5295
      End
      Begin VB.Label lblStep2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting to host site and retrieving version information..."
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Tag             =   "10004"
         Top             =   360
         Width           =   5295
      End
      Begin VB.Label lblYes 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "A new update is available, click Next to download it or click Cancel to exit."
         Height          =   615
         Left            =   240
         TabIndex        =   9
         Tag             =   "10006"
         Top             =   1320
         Visible         =   0   'False
         Width           =   5295
      End
   End
End
Attribute VB_Name = "vcUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sTempFolder As String
Private sConfigFileURL As String
Private sThisVer As String
Private lStep As Long
Private sName As String
Private sWebVer As String
Private sUpdateURL As String
Private sUpdate As String
Private sOutFile As String
Private sSuccessFile As String
Private lMaxBytes As Long
Private m_AutoCheck As Boolean
Private Const sTempFile As String = "_vclu.txt"

Private Declare Function CreateDir Lib "shell32" Alias "SHCreateDirectoryExA" (ByVal hWnd As Long, ByVal pszPath As String, ByVal psa As Any) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Event CloseRequest(DownloadedFile As String)
Public Event UpdateAvailable(IsAvailable As Boolean)

Public Function NextStep()
    cmdNext_Click
End Function

Public Function IsUpdateReady() As Boolean
    If LenB(sSuccessFile) <> 0 Then
        If cmdCancel.Caption = LoadResString(10015) Then
            If cmdCancel.Enabled Then
                IsUpdateReady = True
            End If
        End If
    End If
End Function

Private Function LongDirFix(sTargetString As String, iMax As Integer) As String
Dim iLblLen As Integer
Dim sTempString As String

    sTempString = sTargetString
    iLblLen = iMax

    If Len(sTempString) <= iLblLen Then
        LongDirFix = sTempString
        Exit Function
    End If
            
    LongDirFix = Left$(sTempString, 10) & "..." & Right$(sTempString, iMax - 23)

End Function

Private Function GetTempDir() As String
Dim sBuffer As String * 260
Dim lLength As Long

    lLength = GetTempPath(Len(sBuffer), sBuffer)
    GetTempDir = Left$(sBuffer, lLength)
    
    If Right$(GetTempDir, 1) <> "\" Then
        GetTempDir = GetTempDir & "\"
    End If
    
End Function

Private Sub cmdCancel_Click()
    
    cmdCancel.Enabled = False
    RaiseEvent CloseRequest(sSuccessFile)
End Sub

Private Sub cmdNext_Click()
Dim x As Long
    
    If AutoCheck = False Then
        cmdCancel.SetFocus
    End If
    
    cmdNext.Enabled = False
    lStep = lStep + 1
    
    For x = 0 To fraFrame.UBound
        If x = lStep Then
            fraFrame(x).Visible = True
        Else
            fraFrame(x).Visible = False
        End If
    Next x
    
    Select Case lStep
        Case 1
            If LenB(sConfigFileURL) <> 0 Then
                BeginDownload sConfigFileURL, sTempFolder & sTempFile
            Else
                MsgBox LoadResString(10012), vbCritical
                'lStep = lStep - 1
                cmdCancel_Click
            End If
        Case 2
            
            CreateDir 0&, sUpdate, ByVal 0&
            
            If FileExists(sUpdate & "*.*") Then
                Kill sUpdate & "*.*"
            End If
            
            If FileExists(App.Path & "\" & App.EXEName & ".zip") Then
                Name App.Path & "\" & App.EXEName & ".zip" As App.Path & "\" & App.EXEName & "_.zip"
            End If
            
            lblUpdate.Caption = Replace(lblUpdate.Caption, "$n", vbNewLine & sName)
            lblUpdate2.Caption = lblUpdate2.Caption & vbNewLine & LongDirFix(sUpdate & sOutFile, 70)
            BeginDownload sUpdateURL, sUpdate & sOutFile
            
    End Select
    
End Sub

Private Sub DownloadComplete(SaveFile As String)

    If SaveFile = sTempFolder & sTempFile Then
        
        GetConfigValues
        lblVersion = LoadResString(10013) & sThisVer & LoadResString(10014) & sWebVer
        
        If sThisVer >= sWebVer Then
            
            lblNo.Visible = True
            cmdNext.Enabled = False
            
            If AutoCheck = False Then
                cmdCancel.SetFocus
            End If
            
        Else
            
            lblYes.Visible = True
            cmdNext.Enabled = True
            
            If AutoCheck = False Then
                cmdNext.SetFocus
            End If
            
        End If
        
        RaiseEvent UpdateAvailable(cmdNext.Enabled)
        
    Else
        
        cmdCancel.Caption = LoadResString(10015)
        cmdCancel.SetFocus
        
        If lMaxBytes \ 1024 < 1024 Then
            lblProgressUpdate.Caption = LoadResString(10016) & Format$(lMaxBytes / 1024, "#0.00") & LoadResString(10017)
        Else
            lblProgressUpdate.Caption = LoadResString(10016) & Format$(lMaxBytes / 1048576, "#00.00") & LoadResString(10018)
        End If
        
        sSuccessFile = SaveFile
        
    End If
End Sub

Private Sub DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
    
    If SaveFile = sTempFolder & sTempFile Then
        lblProgress.Caption = CurBytes & " / " & MaxBytes & LoadResString(10019)
    Else
        
        If MaxBytes \ 1024 < 1024 Then
            lblProgressUpdate.Caption = Format$(CurBytes / 1024, "#0.00") & LoadResString(10020) & Format$(MaxBytes / 1024, "#0.00") & LoadResString(10021)
        Else
            lblProgressUpdate.Caption = Format$(CurBytes / 1048576, "#00.00") & LoadResString(10022) & Format$(MaxBytes / 1048576, "#00.00") & LoadResString(10023)
        End If
        
        DoEvents
        lMaxBytes = MaxBytes
        
    End If
    
End Sub

Private Sub LoadResStrings()
Dim ctl As Control
Dim lVal As Long
Dim lLocaleID

    lLocaleID = GetUserDefaultLCID
    
    For Each ctl In UserControl.Controls
        lVal = Val(ctl.Tag)
        If lVal > 0 Then
            ctl.Caption = LoadResString(lVal)
            If lLocaleID = SLOVAK_LOCALE Or lLocaleID = CZECH_LOCALE Then
                ctl.font.Charset = EASTEUROPE_CHARSET
            End If
        End If
    Next
    
End Sub

Private Sub UserControl_Initialize()
    
    LoadResStrings
    
    sTempFolder = GetTempDir
    
    lStep = 0
    sThisVer = App.Major & "." & App.Minor & "." & App.Revision
    sSuccessFile = vbNullString
    
End Sub

Private Sub UserControl_Resize()
    Width = 5970
    Height = 3735
End Sub

Public Property Get AutoCheck() As Boolean
    AutoCheck = m_AutoCheck
End Property

Public Property Let AutoCheck(bool As Boolean)
    m_AutoCheck = bool
    PropertyChanged "AutoCheck"
End Property

Public Property Get ConfigFileURL() As String
    ConfigFileURL = sConfigFileURL
End Property

Public Property Let ConfigFileURL(ps As String)
    sConfigFileURL = ps
    PropertyChanged "ConfigFileURL"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ConfigFileURL", sConfigFileURL, vbNullString
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    sConfigFileURL = PropBag.ReadProperty("ConfigFileURL", vbNullString)
End Sub

Private Sub GetConfigValues()
Const vbPipe As String = "|"
Dim iFileNum As Long
Dim sTrash As String
Dim lRet As Long
Dim sData() As String ', x As Long, sPOS As Long
    
    sName = vbNullString
    sWebVer = vbNullString
    sUpdateURL = vbNullString
    iFileNum = FreeFile
    
    Open sTempFolder & sTempFile For Input As #iFileNum
        If LOF(iFileNum) > 0 Then
            Line Input #iFileNum, sTrash
        End If
    Close #iFileNum
    
    sData = Split(sTrash, vbPipe)
    sName = sData(0)
    sWebVer = sData(1)
    sUpdateURL = sData(2)
    
    If Asc(Right$(sUpdateURL, 1)) = 10 Then
        sUpdateURL = Left$(sUpdateURL, Len(sUpdateURL) - 1)
    End If
    
    lRet = InStrRev(sUpdateURL, "/")
    sOutFile = Right$(sUpdateURL, Len(sUpdateURL) - lRet)
    
    Kill sTempFolder & sTempFile
    sUpdate = App.Path & "\" & LoadResString(10024) & "\"
    Exit Sub
    
ErrorGetConfigValues:
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
Dim f() As Byte, iFileNum As Integer

    If AsyncProp.BytesMax <> 0 Then
        
        iFileNum = FreeFile
        f = AsyncProp.Value
        
        Open AsyncProp.PropertyName For Binary As #iFileNum
            Put #iFileNum, , f
        Close #iFileNum
        
        Call DownloadComplete(AsyncProp.PropertyName)
        
    Else
        ErrorMessage
    End If
    
End Sub

Private Sub ErrorMessage()
    If lStep = 1 Then
        lblProgress.Caption = LoadResString(10025)
    Else
        lblProgressUpdate.Caption = LoadResString(10025)
    End If
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    If AsyncProp.BytesMax <> 0 Then
        Call DownloadProgress(CLng(AsyncProp.BytesRead), CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
    End If
End Sub

Private Sub BeginDownload(URL As String, SaveFile As String)
    
    On Error GoTo ErrorBeginDownload
    AsyncRead URL, vbAsyncTypeByteArray, SaveFile, vbAsyncReadForceUpdate
    Exit Sub
    
ErrorBeginDownload:
    MsgBox LoadResString(10026) & Err.Number & LoadResString(10027) & _
    vbNewLine & LoadResString(10028) & Err.Description & ".", vbCritical
    cmdCancel_Click
End Sub
