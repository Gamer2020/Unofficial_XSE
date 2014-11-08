Attribute VB_Name = "modErrorHandler"
Option Explicit

' Copyright © 2009 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128&
End Type

Private Const sErrorLogFile As String = "\Bug.log"
Private Const sDot As String = "."
Private sSystemInfo As String
Private sAppVersion As String

Private Declare Function GetVersionExA Lib "kernel32" (ByRef lpVersionInfo As OSVERSIONINFO) As Long

Private Sub LogError(ByRef ErrorProcedure As String, ByRef ErrorNumber As Long, ByRef ErrorSource As String, ByRef ErrorDescr As String, ByRef UsersChoice As Long)
Const sSep As String = "|"
Dim iFileNum As Integer
Dim osv As OSVERSIONINFO
Dim sDate As String
Dim sTime As String
Dim sErrorInfo As String
    
    On Error Resume Next
    
    ' Get next available file number
    iFileNum = FreeFile
    
    ' Get date and time in a compacted form
    sDate = Right$("000" & CStr(Year(Now)), 4&) & Right$("00" & CStr(Month(Now)), 2&) & _
        Right$("00" & CStr(Day(Now)), 2&)
    sTime = Right$("00" & CStr(Hour(Now)), 2&) & Right$("00" & CStr(Minute(Now)), 2&) & _
        Right$("00" & CStr(Second(Now)), 2&)
    
    ' Check if the OS info is empty
    If LenB(sSystemInfo) = 0& Then
    
        ' Retrieve the OS info
        osv.dwOSVersionInfoSize = Len(osv)
        GetVersionExA osv
        sSystemInfo = osv.dwMajorVersion & sDot & osv.dwMinorVersion & sDot & _
            osv.dwBuildNumber
        
    End If
        
    ' Check if the app version is empty
    If LenB(sAppVersion) = 0& Then
    
        ' Get the app version
        sAppVersion = App.Major & sDot & App.Minor & sDot & App.Revision
        
    End If
    
    ' Open the error log file and write the collected info
    Open App.Path & sErrorLogFile For Append As #iFileNum
        Print #iFileNum, sDate & sSep & sTime & sSep & _
        sSystemInfo & sSep & App.Title & sSep & sAppVersion & sSep & _
        ErrorProcedure & sSep & ErrorNumber & sSep & ErrorSource & sSep & _
        ErrorDescr
    Close #iFileNum
        
End Sub

Public Function GlobalHandler(ByVal ErrorProcedure As String, ByVal ErrorSource As String) As Long
Dim lErrorNum As Long
Dim sErrorDescr As String
Dim sErrorInfo As String
Dim lRetMsgbox As Long

    On Error Resume Next
    
    ' Get the error info
    lErrorNum = Err.Number
    sErrorDescr = Err.Description
    Err.Clear
    
    ' Make sure the error number is not zero
    If lErrorNum <> 0& Then
    
        ' Display the error message
        lRetMsgbox = MsgBox(sErrorDescr & sDot & vbNewLine & _
            ErrorProcedure & "@" & ErrorSource, _
            vbAbortRetryIgnore + vbDefaultButton2 + vbCritical)
    
        ' Log the error if not in IDE
        If App.LogMode <> 0 Then
            LogError ErrorProcedure, lErrorNum, ErrorSource, sErrorDescr, lRetMsgbox
        End If
        
        ' Set the return value
        GlobalHandler = lRetMsgbox
    
    End If
    
End Function

Public Function Quit()
    On Error Resume Next
    Unload frmMain
    End
End Function
