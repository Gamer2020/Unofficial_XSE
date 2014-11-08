VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Copyright © 2009 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Private Const sMyName As String = "clsDownload"
Private Const lDefaultBufferSize As Long = 8192&

Private Const INTERNET_FLAG_DONT_CACHE As Long = &H4000000
Private Const INTERNET_FLAG_RELOAD As Long = &H80000000
Private Const HTTP_QUERY_CONTENT_LENGTH As Long = &H5&
Private Const HTTP_QUERY_FLAG_NUMBER As Long = &H20000000

Private Declare Function InternetOpenW Lib "wininet" (ByVal lpszAgent As Long, ByVal dwAccessType As Long, ByVal lpszProxyName As Long, ByVal lpszProxyBypass As Long, ByVal dwFlags As Long) As Long
Private Declare Function InternetOpenUrlW Lib "wininet" (ByVal hInternetSession As Long, ByVal lpszUrl As Long, ByVal lpszHeaders As Long, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal dwNumberOfBytesToRead As Long, ByRef lpdwNumberOfBytesRead As Long) As Boolean
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Boolean
Private Declare Function HttpQueryInfoW Lib "wininet" (ByVal hRequest As Long, ByVal dwInfoLevel As Long, ByRef lpvBuffer As Long, ByRef lpdwBufferLength As Long, ByVal lpdwIndex As Long) As Boolean
Private Declare Function DeleteFileW Lib "kernel32" (ByVal lpFileName As Long) As Long

Private m_BufferSize As Long

Public Event DownloadProgress(FileName As String, CurrentBytes As Long, TotalBytes As Long)
Public Event DownloadComplete(FileName As String, TotalBytes As Long)

Public Property Get BufferSize() As Long
    BufferSize = m_BufferSize
End Property

Public Property Let BufferSize(ByVal NewValue As Long)
    
    ' Ensure it's higher than zero
    If NewValue Then
        BufferSize = m_BufferSize
    End If
    
End Property

Public Sub DownloadFile(ByVal URL As String, ByVal FileName As String)
Const sThis = "DownloadFile"
Dim hInternetSession As Long
Dim hUrl As Long
Dim lFileLen As Long
Dim bBuffer() As Byte
Dim lBytesRead As Long
Dim lpBuffer As Long
Dim lCount As Long
Dim iFileNum As Integer
Dim fOkay As Boolean
    
    On Error GoTo LocalHandler
    
    ' Open an Internet session and retrieve its handle
    hInternetSession = InternetOpenW(0&, 0&, 0&, 0&, 0&)
    
    If hInternetSession <> 0& Then
    
        ' Open the file and get its handle
        hUrl = InternetOpenUrlW(hInternetSession, StrPtr(URL), 0&, 0&, INTERNET_FLAG_DONT_CACHE Or INTERNET_FLAG_RELOAD, 0&)
        
        If hUrl <> 0& Then
            
            ' Query the length of the file
            lpBuffer = 4&
            fOkay = HttpQueryInfoW(hUrl, HTTP_QUERY_CONTENT_LENGTH Or HTTP_QUERY_FLAG_NUMBER, lFileLen, lpBuffer, 0&)
            
            ' Make sure the files is not empty
            If fOkay And (lFileLen > 0&) Then
                
                ' Allocate the temp array
                ReDim bBuffer(lFileLen - 1&)
                
                ' Save the buffer pointer
                lpBuffer = VarPtr(bBuffer(0))
                
                ' Check if the buffer is higher than zero
                If m_BufferSize Then
        
                    Do
                        ' Read a chunk of the file
                        fOkay = InternetReadFile(hUrl, lpBuffer + lCount, m_BufferSize, lBytesRead)
                        
                        ' Update the counter
                        lCount = lCount + lBytesRead
                        
                        ' Exit if error or no more data
                        If lBytesRead = 0& Then
                            Exit Do
                        ElseIf fOkay = False Then
                            Exit Do
                        End If
                        
                        RaiseEvent DownloadProgress(FileName, lCount, lFileLen)
                        
                    Loop
                    
                    ' The file was downloaded properly
                    If lCount = lFileLen Then
                        
                        ' Ensure that there is no local file
                        DeleteFileW StrPtr(FileName)
                        
                        ' Get the next free number
                        iFileNum = FreeFile
                        
                        Open FileName For Binary As #iFileNum
                            ' Save the data to the local file
                            Put #iFileNum, , bBuffer
                        Close #iFileNum
                            
                        RaiseEvent DownloadComplete(FileName, lFileLen)
                        
                    Else
                        RaiseEvent DownloadComplete(FileName, 0&)
                    End If
                    
                Else
                    RaiseEvent DownloadComplete(FileName, 0&)
                End If
            
            Else
                RaiseEvent DownloadComplete(FileName, 0&)
            End If
            
            InternetCloseHandle hUrl
            
        Else
            RaiseEvent DownloadComplete(FileName, 0&)
        End If
        
        InternetCloseHandle hInternetSession
        
    Else
        RaiseEvent DownloadComplete(FileName, 0&)
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

Private Sub Class_Initialize()
    m_BufferSize = lDefaultBufferSize
End Sub
