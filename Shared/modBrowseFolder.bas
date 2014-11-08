Attribute VB_Name = "modBrowseFolder"
Option Explicit

' Copyright © 2009 HackMew
' ------------------------------
' Feel free to create derivate works from it, as long as you clearly give me credits of my code and
' make available the source code of derivative programs or programs where you used parts of my code.
' Redistribution is allowed at the same conditions.

Private Const MAX_PATH As Long = 260&

Private Const BIF_RETURNONLYFSDIRS As Long = &H1&
Private Const BIF_NEWDIALOGSTYLE As Long = &H40&
Private Const BIF_NONEWFOLDERBUTTON As Long = &H200&

Private Type BROWSEINFO
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type

Private Declare Function SHBrowseForFolderW Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDListW Lib "shell32" (ByVal pidl As Long, ByVal pszPath As Long) As Long

Public Function BrowseFolder(ByVal hWndOwner As Long, ByRef Title As String) As String
Dim lpIDList As Long
Dim sBuffer As String
Dim BI As BROWSEINFO
    
    ' Initialize the BROWSEINFO structure
    With BI
        .hWndOwner = hWndOwner
        .lpszTitle = StrPtr(Title)
        .ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE Or BIF_NONEWFOLDERBUTTON
    End With

    ' Show the BrowseForFolder dialog
    lpIDList = SHBrowseForFolderW(BI)

    ' Get the selected folder
    If lpIDList Then
        sBuffer = Space$(MAX_PATH)
        SHGetPathFromIDListW lpIDList, StrPtr(sBuffer)
        BrowseFolder = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1&)
    End If
    
End Function


