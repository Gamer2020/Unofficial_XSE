VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Live Update"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmUpdate.frx":000C
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   399
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "10000"
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   345
      Left            =   4680
      TabIndex        =   1
      Tag             =   "10011"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3360
      TabIndex        =   0
      Tag             =   "10010"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Step 1 of 3"
      Height          =   2055
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Tag             =   "10001"
      Top             =   1080
      Width           =   5775
      Begin VB.Label lblNext 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Click Next to continue with the Live Update process."
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Tag             =   "10012"
         Top             =   1320
         Width           =   3750
      End
      Begin VB.Label lblStep1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmUpdate.frx":135C
         Height          =   855
         Left            =   240
         TabIndex        =   3
         Tag             =   "10002"
         Top             =   360
         Width           =   5280
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Step 3 of 3"
      Height          =   2055
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Tag             =   "10007"
      Top             =   1080
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label lblDestination 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destination:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Tag             =   "10009"
         Top             =   720
         Width           =   870
      End
      Begin VB.Label lblArchivePath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  "
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   90
      End
      Begin VB.Label lblProgressUpdate 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000011&
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   5280
      End
      Begin VB.Label lblDownloading 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Downloading update..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Tag             =   "10008"
         Top             =   360
         Width           =   1650
      End
   End
   Begin VB.Frame fraFrame 
      Caption         =   "Step 2 of 3"
      Height          =   2055
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Tag             =   "10003"
      Top             =   1080
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label lblNewestVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  "
         Height          =   195
         Left            =   480
         TabIndex        =   14
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label lblStep2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Connecting to host site and retrieving version information..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Tag             =   "10004"
         Top             =   360
         Width           =   4320
      End
      Begin VB.Label lblProgress 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000011&
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   5280
      End
      Begin VB.Label lblThisVersion 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  "
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   90
      End
      Begin VB.Label lblNo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "No update available. Click Cancel to exit."
         ForeColor       =   &H80000011&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Tag             =   "10005"
         Top             =   1320
         Visible         =   0   'False
         Width           =   5280
      End
      Begin VB.Label lblYes 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "  "
         ForeColor       =   &H80000011&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Tag             =   "10006"
         Top             =   1320
         Visible         =   0   'False
         Width           =   5280
      End
   End
End
Attribute VB_Name = "frmUpdate"
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

Private Const sMyName As String = "frmUpdate"
Private Const MAX_PATH As Long = 260&
Private Const sSpace As String = " "

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Private Const FO_MOVE As Long = &H1&
Private Const FO_COPY As Long = &H2&
Private Const FOF_SILENT As Long = &H4&
Private Const FOF_NOCONFIRMATION As Long = &H10&
Private Const FOF_FILESONLY As Long = &H80&
Private Const FOF_NOCONFIRMMKDIR As Long = &H200&
Private Const FOF_NOERRORUI As Long = &H400&
Private Const FOF_NORECURSION As Long = &H1000&
Private Const FOF_NO_UI As Long = FOF_SILENT Or FOF_NOCONFIRMATION Or FOF_NOCONFIRMMKDIR Or FOF_NOERRORUI

Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function GetTempPathW Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As Long) As Long
Private Declare Function SHCreateDirectoryExW Lib "shell32" (ByVal hWnd As Long, ByVal pszPath As Long, ByVal psa As Long) As Long
Private Declare Function SHFileOperationW Lib "shell32" (ByVal lpFileOp As Long) As Long
Private Declare Function ShellExecuteW Lib "shell32" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Private Declare Function PathCompactPathExW Lib "shlwapi" (ByVal pszOut As Long, ByVal pszSrc As Long, ByVal cchMax As Long, ByVal dwFlags As Long) As Boolean

Private WithEvents Updater As clsDownload
Attribute Updater.VB_VarHelpID = -1
Private fIsUpdateAvailable As Boolean
Private fIsUpdateReady As Boolean
Private fAlreadyUpdating As Boolean
Private sThisVer As String
Private sNewestVer As String
Private sTempFile As String
Private sUpdateFolder As String
Private lCurrentStep As Long
Private sUpdateInfoUrl As String
Private sUpdateUrl As String
Private sArchiveFile As String
Private m_IsUnattended As Boolean

Public Property Get IsUnattended() As Boolean
    IsUnattended = m_IsUnattended
End Property

Public Property Let IsUnattended(ByVal NewValue As Boolean)
    m_IsUnattended = NewValue
End Property

Private Function GetFileTitle(ByRef FileName As String) As String
Dim lPos As Long

    ' Search last backslash
    lPos = InStrRev(FileName, "\")

    ' Ensure there is one
    If lPos Then
        GetFileTitle = Mid$(FileName, lPos + 1&)
    Else
        ' Return the file name as is
        GetFileTitle = FileName
    End If

End Function

Private Function GetUrlFileTitle(ByRef FileUrl As String) As String
Dim lPos As Long

    ' Search last slash
    lPos = InStrRev(FileUrl, "/")

    ' Ensure there is one
    If lPos Then
        GetUrlFileTitle = Mid$(FileUrl, lPos + 1&)
    End If

End Function

Private Function QualifyPath(ByRef Path As String) As String
        
    ' Make sure the path is not empty
    If LenB(Path) Then
    
        ' Check if the last character is a backslash
        If Right$(Path, 1&) <> "\" Then
            
            ' Append a backslash
            QualifyPath = Path & "\"
            
        Else
            ' Leave the path as it was
            QualifyPath = Path
        End If

   End If

End Function

Private Function GetTempPath() As String
Const sThis As String = "GetTempPath"
Dim sBuffer As String
Dim lLength As Long
    
    On Error GoTo LocalHandler
    
    ' Allocate a string buffer
    sBuffer = Space$(MAX_PATH)
    
    ' Get the temp path, retrive the string length
    lLength = GetTempPathW(MAX_PATH, StrPtr(sBuffer))
    
    ' Set the return value
    GetTempPath = QualifyPath(Left$(sBuffer, lLength))
    Exit Function
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
    
End Function

Private Function CompactPath(ByRef Path As String, ByVal MaxChars As Long) As String
Const sThis As String = "CompactPath"
    
    On Error GoTo LocalHandler
    
    ' Check if the path is too long
    If Len(Path) > MaxChars Then
        
        ' Allocate a buffer and compact the path
        CompactPath = Space$(MaxChars + 1&)
        PathCompactPathExW StrPtr(CompactPath), StrPtr(Path), MaxChars + 1&, 0&
        
    Else
        ' Nothing to do
        CompactPath = Path
    End If
    Exit Function
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
    
End Function

Private Function PathExists(Path As String) As Boolean
    
    On Error GoTo Hell
    
    ' Ensure if the path is not empty
    If LenB(Path) Then
        
        ' Check if it's an actual directory
        If LenB(Dir$(Path, vbDirectory)) Then
            PathExists = True
        End If
        
    End If
    
Hell:
End Function

Private Function GetRandomSuffix() As String
    GetRandomSuffix = CStr(Hour(Now)) & CStr(Minute(Now)) & CStr(Second(Now)) & ChrW$(Int(Rnd * 25&) + 98&)
End Function

Private Sub DoUpdate()
Const sThis As String = "DoUpdate"
Const sBatchPrefix As String = "DeleteOld_"
Const sAsterisk As String = "*"
Const sQuote As String = """"
Const sBatExt As String = ".bat"
Const sExeExt As String = ".exe"
Const sNewExt As String = ".new"
Dim iFileNum As Integer
Dim sBatchFile As String
Dim sQualifiedAppPath As String
Dim sAppTitle As String
Dim sAppExeName As String
Dim shfop As SHFILEOPSTRUCT

    On Error GoTo LocalHandler
    
    ' Get the app path and the batch file name
    sQualifiedAppPath = QualifyPath(App.Path)
    sBatchFile = sQualifiedAppPath & sBatchPrefix & GetRandomSuffix & sBatExt
    
    ' Delete all the existing batch files, if any
    If FileExists(sQualifiedAppPath & sBatchPrefix & sAsterisk & sBatExt) Then
        Kill sQualifiedAppPath & sBatchPrefix & sAsterisk & sBatExt
    End If
    
    With shfop
        
        ' Initialize the SHFILEOP structure
        .wFunc = FO_COPY
        .pFrom = sArchiveFile & "\" & sAsterisk & vbNullChar
        .pTo = sUpdateFolder & vbNullChar
        .fFlags = FOF_NO_UI Or FOF_NORECURSION Or FOF_FILESONLY
    
        ' Unzip the files inside the archive
        SHFileOperationW VarPtr(shfop)
        
        ' Get the app title
        sAppTitle = GetFileTitle(sArchiveFile)
        sAppTitle = Left$(sAppTitle, InStr(sAppTitle, ".") - 1&)
        
        ' Append the .new extension to the updated .exe
        If FileExists(sUpdateFolder & sAppTitle & sExeExt) Then
            Name sUpdateFolder & sAppTitle & sExeExt As sUpdateFolder & sAppTitle & sExeExt & sNewExt
        End If
        
        ' Move all the files from the update folder to the app path
        .wFunc = FO_MOVE
        .pFrom = sUpdateFolder & sAsterisk & vbNullChar
        .pTo = sQualifiedAppPath & vbNullChar
        
        SHFileOperationW VarPtr(shfop)
        
        ' Put the archive back in the update folder
        .pFrom = sQualifiedAppPath & sAppTitle & ".zip" & vbNullChar
        .pTo = sUpdateFolder & vbNullChar
    
        SHFileOperationW VarPtr(shfop)
        
    End With
    
    iFileNum = FreeFile
    
    ' Create the batch file
    Open sBatchFile For Output As #iFileNum
        
        sAppExeName = sAppTitle & sExeExt
                         
        Print #iFileNum, ":d" & vbNewLine & _
                            "del /f /q " & sAppExeName & vbNewLine & _
                            "if exist " & sAppExeName & " goto d" & vbNewLine & _
                            ":r" & vbNewLine & _
                            "ren " & sAppExeName & sNewExt & sSpace & sAppExeName & vbNewLine & _
                            "if exist " & sAppExeName & sNewExt & " goto r" & vbNewLine & _
                            "if exist " & sAppExeName & " start " & sAppExeName & vbNewLine & _
                            "del /f /q " & sBatchPrefix & sAsterisk & sBatExt & vbNewLine & _
                            "exit";
                  
    Close #iFileNum
    
    ' Run the batch file, make it hidden
    ShellExecuteW Me.hWnd, StrPtr("open"), StrPtr(sBatchFile), 0&, StrPtr(sQualifiedAppPath), vbHide
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

Private Sub cmdCancel_Click()
Const sThis As String = "cmdCancel_Click"

    On Error GoTo LocalHandler
    
    ' If the button is set as Finish
    If lCurrentStep > 1& Then
    
        ' Disable the button to prevent the user clicking it again
        cmdCancel.Enabled = False
        cmdCancel.Refresh
        DoEvents
        
    End If
    
    ' Check if the update was downloaded
    If fIsUpdateReady Then
        
        ' Start the program update
        fAlreadyUpdating = True
        Screen.MousePointer = vbHourglass
        DoUpdate
        
        ' Restore the mouse pointer
        Screen.MousePointer = vbDefault
        
        ' Exit
        Unload frmMain
        End
        
    Else
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

Public Sub NextStep()
    cmdNext_Click
End Sub

Private Sub cmdNext_Click()
Const sThis As String = "cmdNext_Click"
Const sAsterisk As String = "*"
Dim i As Long
    
    On Error GoTo LocalHandler
    
    ' Don't focus the Cancel button when in silent mode
    If m_IsUnattended = False Then
        cmdCancel.SetFocus
    End If
    
    ' Disable the button
    cmdNext.Enabled = False
    
    ' Increase step
    lCurrentStep = lCurrentStep + 1&
    
    ' Show the right frame
    For i = fraFrame.LBound To fraFrame.UBound
        fraFrame(i).Visible = (i = lCurrentStep)
    Next i
    
    If lCurrentStep = 1& Then
        
        ' Download the update info
        Updater.DownloadFile sUpdateInfoUrl, sTempFile
        
        ' Check if running unattended
        If m_IsUnattended Then
        
            ' Check if there are some updates
            If fIsUpdateAvailable Then
                
                ' Enable the form show timer
                frmMain.tmrShowUpdate.Enabled = True
                
            Else
                ' Exit
                Unload Me
                Exit Sub
            End If
        
        End If
        
    Else
            
        ' If the update folder doesn't exist
        If PathExists(sUpdateFolder) = False Then
            
            ' Create it
            SHCreateDirectoryExW 0&, StrPtr(sUpdateFolder), 0&
            
        Else
            
            ' The folder already exists, clean it
            If FileExists(sUpdateFolder & sAsterisk) Then
                Kill sUpdateFolder & sAsterisk
            End If
        
        End If
        
        ' Display the compacted archive path
        lblArchivePath.Caption = CompactPath(sArchiveFile, 70&)
        
        ' Download the update
        Updater.DownloadFile sUpdateUrl, sArchiveFile
            
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    ' Prevent form close when the update is ready
    If fAlreadyUpdating = False Then
        Cancel = fIsUpdateReady
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Const sThis = "Form_KeyPress"
    
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

    On Error GoTo LocalHandler
    
    ' Localize the form
    Localize Me
    
    ' Initialize the PRNG
    Randomize
    
    ' Create an updater
    Set Updater = New clsDownload
    
    ' Set basic update data
    sThisVer = App.Major & sDot & App.Minor & sDot & App.Revision
    sTempFile = GetTempPath & "_upd" & GetRandomSuffix & ".tmp"
    sUpdateInfoUrl = "http://www.andreasartori.net/hackmew/updates/" & App.Title & ".txt"
    sUpdateFolder = QualifyPath(App.Path) & LoadString(ID_UPDATEFOLDER) & "\"
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
Const sThis As String = "Form_Unload"

    On Error GoTo LocalHandler
    
    ' Free the memory associated with the form
    Set frmUpdate = Nothing
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

Private Sub Updater_DownloadComplete(FileName As String, TotalBytes As Long)
Const sThis As String = "Updater_DownloadComplete"
Const sNullVersion As String = "0.0.0"
Const sFormat As String = "#0.00"
Dim iFileNum As Integer
Dim sRaw As String
Dim sData() As String

    On Error GoTo LocalHandler
    
    ' Check the currect step
    If lCurrentStep = 1& Then
    
        ' Ensure something was actually downloaded
        If TotalBytes Then
        
            iFileNum = FreeFile
        
            ' Open the update info file
            Open sTempFile For Input As #iFileNum
            
                ' Not truly needed, but better be sure
                If LOF(iFileNum) Then
                    
                    ' Check if we can get a line
                    If EOF(iFileNum) = False Then
                        Line Input #iFileNum, sRaw
                    End If
                    
                End If
                
            Close #iFileNum
            
            ' Split the line
            sData = Split(sRaw, "|")
            
            ' Delete the info file
            DeleteFileW StrPtr(sTempFile)
            
            ' The UBound must be 2 for a valid file
            ' AppName|Version|UpdateUrl
            If UBound(sData) = 2& Then
                
                ' In the unlikely case the file just *seemed*
                ' valid, check if the version is a number
                If IsNumeric(Replace(sData(1), ".", vbNullString)) Then
                
                    ' Retrieve the info
                    sNewestVer = sData(1)
                    sUpdateUrl = sData(2)
                    
                    ' Ensure the url is pointing to a file
                    If InStr(sUpdateUrl, ".zip") Then
                    
                        ' Set the archive file and update the progress
                        sArchiveFile = sUpdateFolder & GetUrlFileTitle(sUpdateUrl)
                        lblProgress.Caption = TotalBytes & " / " & TotalBytes & sSpace & LoadString(ID_BYTESRETRIEVED)
                    
                    Else
                        ' Something went wrong, can't update
                        sNewestVer = sNullVersion
                        lblProgress.Caption = LoadString(ID_CONNECTIONERROR)
                    End If
                
                Else
                    ' Something went wrong, can't update
                    sNewestVer = sNullVersion
                    lblProgress.Caption = LoadString(ID_CONNECTIONERROR)
                End If
                
            Else
                ' Not a proper file
                sNewestVer = sNullVersion
                lblProgress.Caption = LoadString(ID_CONNECTIONERROR)
            End If
            
            ' Is the current version the newest one?
            If sThisVer >= sNewestVer Then
                
                ' Seems so. Tell the user there are no updates
                lblNo.Visible = True
                lblThisVersion.FontBold = True
                cmdNext.Enabled = False
                
                ' Set the focus to the Cancel button
                ' if the form is shown
                If m_IsUnattended = False Then
                    cmdCancel.SetFocus
                End If
                
            Else
            
                ' There's a new update, yay!
                lblYes.Visible = True
                lblNewestVersion.FontBold = True
                cmdNext.Enabled = True
                
                ' Don't focus if running IsUnattended
                If m_IsUnattended = False Then
                    cmdNext.SetFocus
                End If
                
                ' Remember there's an update for later
                fIsUpdateAvailable = True
                
            End If
            
            ' Set the version labels
            lblThisVersion.Caption = LoadString(ID_YOURVERSION) & sSpace & sThisVer
            lblNewestVersion.Left = lblThisVersion.Left + lblThisVersion.Width + (30& * Screen.TwipsPerPixelX)
            lblNewestVersion = LoadString(ID_NEWESTVERSION) & sSpace & sNewestVer
            
        Else
            
            ' Nothing was downloaded
            ' If running IsUnattended, exit
            If m_IsUnattended Then
                Unload Me
                Exit Sub
            End If
            
            ' Tell the user there was a problem
            lblProgress.Caption = LoadString(ID_CONNECTIONERROR)

        End If
        
    Else
        
        ' Make sure something was downloaded
        If TotalBytes Then
            
            ' Update the Cancel button, focus it
            cmdCancel.Caption = LoadString(ID_FINISH)
            cmdCancel.SetFocus
            
            ' Update the progress label
            If (TotalBytes \ 1024&) < 1024& Then
                lblProgressUpdate.Caption = LoadString(ID_DOWNLOADOF) & sSpace & Format$(TotalBytes / 1024&, sFormat) & _
                    sSpace & LoadString(ID_KILOBYTE) & sSpace & LoadString(ID_COMPLETE)
            Else
                lblProgressUpdate.Caption = LoadString(ID_DOWNLOADOF) & sSpace & Format$(TotalBytes / 1048576, sFormat) & _
                    sSpace & LoadString(ID_MEGABYTE) & sSpace & LoadString(ID_COMPLETE)
            End If
            
            ' Mark the fact the udpate is ready
            fIsUpdateReady = True
        
        Else
            ' Nothing downloaded, sadly
            lblProgressUpdate.Caption = LoadString(ID_CONNECTIONERROR)
        End If
        
        ' Refresh the label
        lblProgressUpdate.Refresh
        
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

Private Sub Updater_DownloadProgress(FileName As String, CurrentBytes As Long, TotalBytes As Long)
Const sThis As String = "Updater_DownloadProgress"
Const lKiloByte As Long = 1024&
Const lMegaByte As Long = 1048576
Const sFormat As String = "#0.00"

    On Error GoTo LocalHandler
    
    ' Make sure we're in the second step
    If lCurrentStep = 2& Then
        
        ' Update the progress label
        If (TotalBytes \ 1024&) < 1024& Then
            lblProgressUpdate.Caption = Format$(CurrentBytes / lKiloByte, sFormat) & " / " & _
                    Format$(TotalBytes / lKiloByte, sFormat) & sSpace & LoadString(ID_KILOBYTE) & sSpace & LoadString(ID_DOWNLOADED)
        Else
            lblProgressUpdate.Caption = Format$(CurrentBytes / lMegaByte, sFormat) & " / " & _
                    Format$(TotalBytes / lMegaByte, sFormat) & sSpace & LoadString(ID_KILOBYTE) & sSpace & LoadString(ID_DOWNLOADED)
        End If
        
        ' Allow the user to cancel while downloading
        ' and at the same time refresh the label
        lblProgressUpdate.Refresh
        DoEvents
        
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
