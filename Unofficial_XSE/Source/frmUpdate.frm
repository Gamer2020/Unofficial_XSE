VERSION 5.00
Begin VB.Form frmUpdate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Live Update"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5970
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
   ScaleHeight     =   244
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "10000"
   Begin VB.FileListBox filUpdatedFiles 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin eXtremeScriptEditor.vcUpdate vcLiveUpdate 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   6588
      ConfigFileURL   =   "http://www.andreasartori.net/hackmew/updates/XSE.txt"
   End
End
Attribute VB_Name = "frmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private AlreadyUpdating As Boolean
Private sFileArray() As String

Private Sub Extract(ByVal sZipFile As String, Optional ByVal sTargetDir As String = vbNullString)
' Function to extract all files from a compressed "folder"
' (ZIP, CAB, etc.) using the Shell Folders' CopyHere method
' (http://msdn2.microsoft.com/en-us/library/ms723207.aspx).
' All files and folders will be extracted from the ZIP file.
' A progress bar will be displayed, and the user will be
' prompted to confirm file overwrites if necessary.
'
' Note:
' This function can also be used to copy "normal" folders,
' if a progress bar and confirmation dialog(s) are required:
' just use a folder path for the "sZipFile" argument.
'
' Arguments:
' sZipFile    [string]  the (path and) file name of the ZIP file
' sTargetDir  [string]  the path of the (existing) destination folder
'
' Based on an article by Gerald Gibson Jr.:
' http://www.codeproject.com/csharp/decompresswinshellapics.asp
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com

Dim objShell As Object
Dim i As Integer

    ' Create the required Shell objects
    Set objShell = CreateObject("Shell.Application")
    
    If LenB(sTargetDir) = 0 Then
        sTargetDir = GetPath(sZipFile)
    End If

    ' UnZIP the files
    objShell.NameSpace((sTargetDir)).CopyHere objShell.NameSpace((sZipFile)).Items()

    ReDim sFileArray(objShell.NameSpace((sZipFile)).Items.Count - 1) As String
    filUpdatedFiles.Path = sTargetDir
    
    For i = 0 To filUpdatedFiles.ListCount - 2
        If InStrB(sFileArray(i), ".zip") = 0 Then
            sFileArray(i) = filUpdatedFiles.List(i)
            If InStrB(1, sFileArray(i), ".exe", vbBinaryCompare) <> 0 Then
                sFileArray(i) = sFileArray(i) & ".new"
            End If
        End If
    Next i
        
    ' Release the object
    Set objShell = Nothing
    
End Sub

Private Sub DeleteOld()
Dim iFileNum As Integer
Dim sBatchName As String
Dim i As Integer
    
    DeleteFile App.Path & "\DeleteOld_*.bat"
    
    Randomize
    iFileNum = FreeFile
    sBatchName = "DeleteOld_" & CStr(Hour(Now)) & CStr(Minute(Now)) & CStr(Second(Now)) & ChrW$(Int(Rnd * 25) + 97) & ".bat"
    
    'create the batch file in the same directory as the old and new versions to make this batch smaller
    Open App.Path & "\" & sBatchName For Output As #iFileNum 'create the batch file
    
        'open the created batch file and print some commands into it, batch file will look like this
            
        '@echo off
        ':s
        'del "(this is the app exe name, we use this incase the user changed the exe name)"   <note: the quotation marks throughout this batch file are nesasary incase your exe name contains spaces>
        'if exist "(app name again here)" goto s   <so if its not deleted yet, go back to :S and read on>
        ':d
        'ren "Update ExampleNEW.exe" "Update Example.exe"   <use the batch to change the new version into the same name as the old version>
        'if exist "Update ExampleNEW.exe" goto d   <same as three lines above>
        'Update Example   <run the new version, name is now the same as old version>
        'del DeleteOld.bat   <delete this batch file>
        
        For i = LBound(sFileArray) To UBound(sFileArray)
            Print #iFileNum, "xcopy /y /c " & ChrW$(34) & App.Path & "\" & LoadResString(10024) & "\" & sFileArray(i) & ChrW$(34) & " " & ChrW$(34) & App.Path & ChrW$(34)
            Print #iFileNum, "del /q " & ChrW$(34) & App.Path & "\" & LoadResString(10024) & "\" & sFileArray(i) & ChrW$(34)
        Next i
                         
        Print #iFileNum, ":d" & vbNewLine & _
                         "if exist " & ChrW$(34) & App.EXEName & ".exe" & ChrW$(34) & " del /q " & ChrW$(34) & App.EXEName & ".exe" & ChrW$(34) & vbNewLine & _
                         "if not exist " & ChrW$(34) & App.EXEName & ".exe" & ChrW$(34) & " goto r" & vbNewLine & _
                         "if exist " & ChrW$(34) & App.EXEName & ".exe" & ChrW$(34) & " goto d" & vbNewLine & _
                         ":r" & vbNewLine & _
                         "ren " & ChrW$(34) & App.EXEName & ".exe.new" & ChrW$(34) & " " & ChrW$(34) & App.EXEName & ".exe" & ChrW$(34) & vbNewLine & _
                         ChrW$(34) & App.EXEName & ".exe" & ChrW$(34) & vbNewLine & _
                         "if exist " & ChrW$(34) & App.EXEName & ".exe.new" & ChrW$(34) & " goto r" & vbNewLine & _
                         "del /q DeleteOld_*.bat" & vbNewLine & _
                         "exit"
                  
    Close #iFileNum
    
    'run the batch file, make it run hidden
    'Shell "DeleteOld.bat", vbHide
   
    ShellExecute Me.hWnd, "open", sBatchName, vbNullString, App.Path, vbHide
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not AlreadyUpdating Then
        Cancel = vcLiveUpdate.IsUpdateReady
    End If
End Sub

Private Sub vcLiveUpdate_CloseRequest(DownloadedFile As String)
    
    If LenB(DownloadedFile) <> 0 Then 'they did download an update
        'ExecuteLink DownloadedFile
        AlreadyUpdating = True
        Extract DownloadedFile
        Name Replace(GetPath(DownloadedFile) & "\" & App.EXEName & ".exe", "\\", "\") As Replace(GetPath(DownloadedFile) & "\" & App.EXEName & ".exe.new", "\\", "\")
        If App.LogMode <> 0 Then DeleteOld
        Unload frmMain
        End
        Exit Sub
    End If
    
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

Private Sub vcLiveUpdate_UpdateAvailable(IsAvailable As Boolean)
    
    If vcLiveUpdate.AutoCheck = True Then
        
        'frmMain.mnuCheckNow.Enabled = True
        
        'If IsAvailable = True Then
        '    Show2 frmUpdate, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
        'Else
        '    Unload Me
        'End If
        
    End If
    
End Sub
