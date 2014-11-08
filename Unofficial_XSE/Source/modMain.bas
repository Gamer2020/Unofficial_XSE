Attribute VB_Name = "modMain"
Option Explicit

Public Const MaxTabLimit As Integer = 20
Public Const MaxUndoSize As Integer = 100

Private Type tPanel
    PanelText As String
    ToolTipTxt As String
    ClientWidth As Long
    pEnabled As Boolean
End Type

Public Document(0 To MaxTabLimit) As frmRubIDE
Public lDocCounter As Long
Public Const CaptionBase As String = "Script "
Public Const vbSpace As String = " "
Public Panels(0 To 7) As tPanel

Public IsLoading As Boolean
Public sEmulatorPath As String

Public Const GWL_STYLE = (-16)
'Public Const GWL_WNDPROC = (-4)

Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Function GetInputState Lib "user32" () As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
'Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Const WM_SETREDRAW = &HB&
Public Const WM_SETTEXT = &HC&

Public Const WM_CUT = &H300&
Public Const WM_COPY = &H301&
Public Const WM_PASTE = &H302&

Public Const EM_GETSEL = &HB0&
Public Const EM_SETSEL = &HB1&
Public Const EM_SCROLLCARET = &HB7&
Public Const EM_GETLINECOUNT = &HBA&
Public Const EM_LINEINDEX = &HBB&
Public Const EM_REPLACESEL = &HC2&
Public Const EM_LINEFROMCHAR = &HC9&

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    x As Long
    Y As Long
End Type

Public CustomColors() As Byte

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Type tagInitCommonControlsEx
   lSize As Long
   lICC As Long
End Type

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function InitCommonControlsEx Lib "COMCTL32.DLL" (iccex As tagInitCommonControlsEx) As Boolean

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXSMICON = 49
Private Const SM_CYSMICON = 50

Private Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Private Const LR_SHARED = &H8000&
Private Const IMAGE_ICON = 1

Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SETICON = &H80
Private Const ICON_SMALL = 0
Private Const ICON_BIG = 1

Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4

'Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Public Const EASTEUROPE_CHARSET = 238
Public Const SLOVAK_LOCALE = &H41B
Public Const CZECH_LOCALE = &H405

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMillisecond As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Private sWindowTitle As String
Private sNewText As String
Private Success As Boolean

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Sub RtlMoveMemory Lib "kernel32" (pDest As Any, pSource As Any, ByVal cb As Long)
Public Declare Sub RtlFillMemory Lib "kernel32.dll" (Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Public Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

' Windows type used to call the Net API
Private Type USER_INFO_10
   User10_Name As Long
   User10_Comment As Long
   User10_User_Comment As Long
   User10_Full_Name As Long
End Type

' Private type to hold the actual strings displayed
Private Type USER_INFO
   name As String
   FullName As String
   Comment As String
   UserComment As String
End Type

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function NetUserGetInfo Lib "netapi32" (lpServer As Byte, Username As Byte, ByVal Level As Long, lpBuffer As Long) As Long
Private Declare Function NetApiBufferFree Lib "netapi32" (ByVal Buffer As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

Private Const PRINTER_ENUM_CONNECTIONS = &H4
Private Const PRINTER_ENUM_LOCAL = &H2

Private Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal flags As Long, ByVal name As String, ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long

Public m_hMod As Long

Public Sub Main()
Const ICC_USEREX_CLASSES = &H200
Dim iccex As tagInitCommonControlsEx

   On Error Resume Next
   
    iccex.lSize = LenB(iccex)
    iccex.lICC = ICC_USEREX_CLASSES
    
    m_hMod = LoadLibrary("shell32.dll")
    InitCommonControlsEx iccex
   
   On Error GoTo Hell
   Load frmMain
   Exit Sub
   
Hell:
    MsgBox LoadResString(10026) & Err.Number & ": " & Err.Description & ".", vbCritical
End Sub

Public Sub SetIcon(ByVal hWnd As Long, ByVal sIconResName As String, Optional ByVal bSetAsAppIcon As Boolean = True)
Dim lhWndTop As Long
Dim lhWnd As Long
Dim cx As Long
Dim cy As Long
Dim hIconLarge As Long
Dim hIconSmall As Long

    If (bSetAsAppIcon) Then
        ' Find VB's hidden parent window:
        lhWnd = hWnd
        lhWndTop = lhWnd
        Do While Not (lhWnd = 0)
            lhWnd = GetWindow(lhWnd, GW_OWNER)
            If Not (lhWnd = 0) Then
                lhWndTop = lhWnd
            End If
        Loop
    End If
    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
    hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
    If (bSetAsAppIcon) Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
    End If
    SendMessageLong hWnd, WM_SETICON, ICON_BIG, hIconLarge
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
    If (bSetAsAppIcon) Then
        SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
    End If
    SendMessageLong hWnd, WM_SETICON, ICON_SMALL, hIconSmall

End Sub

Public Sub Localize(frm As Form)
Dim ctl As Control
Dim lVal As Long
Dim lLocaleID As Long
    
    ' set the form's caption
    If Val(frm.Tag) > 0 Then
        frm.Caption = LoadResString(Val(frm.Tag))
    End If
    
    lLocaleID = GetUserDefaultLCID
    
    If lLocaleID = SLOVAK_LOCALE Then
        frm.font.Charset = EASTEUROPE_CHARSET
    ElseIf lLocaleID = CZECH_LOCALE Then
        frm.font.Charset = EASTEUROPE_CHARSET
    End If
    
    For Each ctl In frm.Controls
        
        If TypeName(ctl) <> "Menu" Then
            lVal = Val(ctl.Tag)
        Else
            lVal = Val(ctl.HelpContextID)
        End If
        
        If lVal > 0 Then
            
            ctl.Caption = LoadResString(lVal)
            
            If lLocaleID = SLOVAK_LOCALE Then
                ctl.font.Charset = EASTEUROPE_CHARSET
            ElseIf lLocaleID = CZECH_LOCALE Then
                ctl.font.Charset = EASTEUROPE_CHARSET
            End If
        
        End If
        
    Next

End Sub

Public Sub LockUpdate(hWnd As Long)
    SendMessage hWnd, WM_SETREDRAW, False, ByVal 0
End Sub

Public Sub UnlockUpdate(hWnd As Long, Optional ToRedraw As Boolean = True)
    SendMessage hWnd, WM_SETREDRAW, True, ByVal 0
    If ToRedraw = True Then
        Redraw hWnd
    End If
End Sub

Public Sub Redraw(hWnd As Long)
Const RDW_INVALIDATE = &H1
Const RDW_INTERNALPAINT = &H2
'Const RDW_NOERASE = &H20
Const RDW_ALLCHILDREN = &H80
Const RDW_UPDATENOW = &H100
'Const RDW_NOFRAME = &H800

   RedrawWindow hWnd, ByVal 0&, 0, RDW_ALLCHILDREN Or RDW_INVALIDATE Or RDW_INTERNALPAINT Or RDW_UPDATENOW
   
End Sub

Public Sub SetStatusText(sText As String)
    frmMain.tmrRestore.Enabled = False
    frmMain.StatusBar.PanelCaption(1) = sText
    frmMain.tmrRestore.Enabled = True
End Sub

'Private Sub CopyProperties(frmDest As frmRubIDE, frmSource As frmRubIDE)
'Dim i As Byte
'Dim sTemp As String
    
'    frmDest.IgnoreChanges = False
'
'    frmDest.Caption = frmSource.Caption
'    'frmDest.txtCode.Text = frmSource.txtCode.Text
'
'    sTemp = frmSource.txtCode.text
'    SendMessage frmDest.txtCode.hWnd, WM_SETTEXT, 0&, ByVal sTemp
'    sTemp = vbNullString
'
'    frmSource.GetCount frmSource.txtCode
'    frmDest.SetCaretPos frmSource.ActualLine, frmSource.ActualColumn
'
'    frmDest.LoadedFile = frmSource.LoadedFile
'    frmDest.txtOffset.Enabled = frmSource.txtOffset.Enabled
'
'    If frmDest.txtOffset.Enabled Then
'        frmDest.txtOffset.text = frmSource.txtOffset.text
'    End If
'
'    frmDest.CanUndo = frmSource.CanUndo
'    frmDest.CanRedo = frmSource.CanRedo
'    frmDest.StackIndex = frmSource.StackIndex
'    frmDest.MaxRedo = frmSource.MaxRedo
'
'    frmDest.EraseStack
'
'    For i = 1 To frmSource.colStack.Count
'        frmDest.colStack.Add frmSource.colStack.Item(i)
'        frmDest.colLine.Add frmSource.colLine.Item(i)
'        frmDest.colCol.Add frmSource.colCol.Item(i)
'    Next i
'
'    frmDest.txtCode_Change
'
'End Sub

Public Function FileLength(sPathName As String) As Long
    If FileExists(sPathName) Then
        FileLength = FileLen(sPathName)
    Else
        FileLength = -1
    End If
End Function

Public Function FileExists(sFilePath As String) As Boolean
    
    On Error GoTo Hell

    If LenB(sFilePath) <> 0 Then
        If LenB(Dir$(sFilePath)) Then
            FileExists = True
        End If
    End If
    
Hell:
End Function

Public Function IsOpen(sFormName As String) As Boolean
Dim i As Long
    
    For i = 0 To Forms.Count - 1
        If InStrB(1, Forms(i).name, sFormName, vbBinaryCompare) <> 0 Then
            IsOpen = True
            Exit Function
        End If
    Next i
    
End Function

Public Function IsReadOnly(sFilePath As String)
    If FileExists(sFilePath) Then
        IsReadOnly = GetAttr(sFilePath) And vbReadOnly
    End If
End Function

Public Sub MakeWritable(sFilePath As String)
    If IsReadOnly(sFilePath) Then
        SetAttr sFilePath, GetAttr(sFilePath) Xor vbReadOnly
    End If
End Sub

Sub SetTopmostWindow(ByVal hWnd As Long, Optional TopMost As Boolean = True, Optional NoActivate As Boolean = False)
Const HWND_NOTOPMOST = -2
Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10

    If TopMost = True Then
        If NoActivate = False Then
            SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
        Else
            SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
        End If
    Else
        SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
    
End Sub

Public Function GetFileName(sFile As String) As String
Dim lPos As Long

  ' search last backslash
    lPos = InStrRev(sFile, "\")

    If lPos <> 0 Then
        GetFileName = Mid$(sFile, lPos + 1)
    Else
        GetFileName = sFile
    End If

End Function

Public Function GetExt(sFile As String) As String
Dim lPos As Long

    ' search last dot
    lPos = InStrRev(sFile, ".")

    If lPos > 0 Then
        If InStr(lPos + 1, sFile, "\", vbBinaryCompare) = 0 Then
            GetExt = LCase$(Mid$(sFile, lPos + 1))
        End If
    End If
    
End Function

Public Function GetPath(sFile As String) As String
Dim lPos As Long
  
  ' search last backslash
  lPos = InStrRev(sFile, "\")
  
  If lPos <> 0 Then
    GetPath = Left$(sFile, lPos)
  Else
    GetPath = sFile
  End If
  
End Function

Public Sub GotoLine(ByVal lLine As Long)
Dim lTemp As Long
    
    lTemp = SendMessage(Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_LINEINDEX, lLine - 1, ByVal 0&)
    SendMessage Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_SETSEL, lTemp, ByVal lTemp
    SendMessage Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, EM_SCROLLCARET, 0&, ByVal 0
    
End Sub

Public Function GetTextBoxLine(hWnd As Long, Optional lLineNumber As Long = -1) As String
Const EM_GETLINE = &HC4
Dim lLine As Long
Dim lLen As Long
Dim sBuffer As String * 1024
    
    If lLineNumber = -1 Then
        lLine = SendMessage(hWnd, EM_LINEFROMCHAR, -1, ByVal 0&) 'get current line
    ElseIf lLineNumber > -1 Then
        lLine = lLineNumber
    End If
    
    Mid$(sBuffer, 1, 4) = ChrW$(Len(sBuffer) And &HFF&) & ChrW$(Len(sBuffer) \ &H100&)
    lLen = SendMessageStr(hWnd, EM_GETLINE, lLine, sBuffer) 'text of line saved to bBuffer
        
    GetTextBoxLine = LeftB$(sBuffer, lLen * 2)

End Function

Private Function GetPointerToByteStringW(lpString As Long) As String
Dim buff() As Byte
Dim nSize As Long
   
    If lpString Then
       
        ' It's Unicode, so multiply by 2
        nSize = lstrlenW(lpString) * 2
        
        If nSize Then
            ReDim buff(0 To (nSize - 1)) As Byte
            RtlMoveMemory buff(0), ByVal lpString, nSize
            GetPointerToByteStringW = buff
        End If
      
    End If
   
End Function

Private Function GetUserNetworkInfo(bServername() As Byte, bUsername() As Byte) As USER_INFO
Dim usrapi As USER_INFO_10
Dim buff As Long
   
    If NetUserGetInfo(bServername(0), bUsername(0), 10, buff) = 0 Then
     
        ' Copy the data from buff into the
        ' API user_10 structure
        RtlMoveMemory usrapi, ByVal buff, Len(usrapi)
        
        ' Extract each member and return
        ' as members of the UDT
        GetUserNetworkInfo.FullName = GetPointerToByteStringW(usrapi.User10_Full_Name)
        
        NetApiBufferFree buff
   
    End If
   
End Function

Public Function Username() As String
Dim bUsername() As Byte
Dim bServername() As Byte
Dim tmp As String
Dim lLen As Long
Dim sUserName As String
    
    ' Set the max buffer length and fill it
    lLen = 256
    sUserName = Space$(lLen)
    
    ' Put the username into the stored buffer
    GetUserName sUserName, lLen
    
    ' Strip any unuseful char
    sUserName = Left$(sUserName, lLen - 1)
    
    ' Assign the clean username to a temp array
    bUsername = sUserName & vbNullChar
    
    ' Prepare a new buffer
    lLen = 16
    tmp = Space$(lLen)
    
    ' Store the computer name into the buffer
    GetComputerName tmp, lLen
    
    ' Trim unuseful spaces
    tmp = Left$(tmp, lLen)
    
    ' Assure the server string is properly formatted
    If LenB(tmp) <> 0 Then
        
        If InStrB(tmp, "\\") Then
            bServername = tmp & vbNullChar
        Else
            bServername = "\\" & tmp & vbNullChar
        End If
        
        ' Get full name
        Username = GetUserNetworkInfo(bServername(), bUsername()).FullName
        
        ' If it's empry, use the default username
        If LenB(Username) = 0 Then
            Username = sUserName
        End If
        
    Else
        Username = sUserName
    End If
    
End Function

Public Function InStrCount(ByRef text As String, ByRef Find As String, Optional ByVal Start As Long = 1, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long
' by Jost Schwider, jost@schwider.de, 20010912, rev 001 20011121
Const MODEMARGIN = 8
Dim TextAsc() As Integer
Dim TextData As Long
Dim TextPtr As Long
Dim FindAsc(0 To MODEMARGIN) As Integer
Dim FindLen As Long
Dim FindChar1 As Integer
Dim FindChar2 As Integer
Dim i As Long

Begin:

  If Compare = vbBinaryCompare Then
    FindLen = Len(Find)
    If FindLen Then
      'Ersten Treffer bestimmen:
      If Start < 2 Then
        Start = InStrB(text, Find)
      Else
        Start = InStrB(Start + Start - 1, text, Find)
      End If

      If Start Then
        InStrCount = 1
        If FindLen <= MODEMARGIN Then

          If TextPtr = 0 Then
            'TextAsc-Array vorbereiten:
            ReDim TextAsc(1 To 1)
            TextData = VarPtr(TextAsc(1))
            RtlMoveMemory TextPtr, ByVal ArrPtr(TextAsc), 4
            TextPtr = TextPtr + 12
          End If

          'TextAsc-Array initialisieren:
          RtlMoveMemory ByVal TextPtr, ByVal VarPtr(text), 4 'pvData
          RtlMoveMemory ByVal TextPtr + 4, Len(text), 4      'nElements

          Select Case FindLen
          Case 1

            'Das Zeichen buffern:
            FindChar1 = AscW(Find)

            'Zählen:
            For Start = Start \ 2 + 2 To Len(text)
              If TextAsc(Start) = FindChar1 Then InStrCount = InStrCount + 1
            Next Start

          Case 2

            'Beide Zeichen buffern:
            FindChar1 = AscW(Find)
            FindChar2 = AscW(Right$(Find, 1))

            'Zählen:
            For Start = Start \ 2 + 3 To Len(text) - 1
              If TextAsc(Start) = FindChar1 Then
                If TextAsc(Start + 1) = FindChar2 Then
                  InStrCount = InStrCount + 1
                  Start = Start + 1
                End If
              End If
            Next Start

          Case Else

            'FindAsc-Array füllen:
            RtlMoveMemory ByVal VarPtr(FindAsc(0)), ByVal StrPtr(Find), FindLen + FindLen
            FindLen = FindLen - 1

            'Die ersten beiden Zeichen buffern:
            FindChar1 = FindAsc(0)
            FindChar2 = FindAsc(1)

            'Zählen:
            For Start = Start \ 2 + 2 + FindLen To Len(text) - FindLen
              If TextAsc(Start) = FindChar1 Then
                If TextAsc(Start + 1) = FindChar2 Then
                  For i = 2 To FindLen
                    If TextAsc(Start + i) <> FindAsc(i) Then Exit For
                  Next i
                  If i > FindLen Then
                    InStrCount = InStrCount + 1
                    Start = Start + FindLen
                  End If
                End If
              End If
            Next Start

          End Select

          'TextAsc-Array restaurieren:
          RtlMoveMemory ByVal TextPtr, TextData, 4 'pvData
          RtlMoveMemory ByVal TextPtr + 4, 1&, 4   'nElements

        Else

          'Konventionell Zählen:
          FindLen = FindLen + FindLen
          Start = InStrB(Start + FindLen, text, Find)
          Do While Start
            InStrCount = InStrCount + 1
            Start = InStrB(Start + FindLen, text, Find)
          Loop

        End If 'FindLen <= MODEMARGIN
      End If 'Start
    End If 'FindLen
  Else
    'Groß-/Kleinschreibung ignorieren:
    text = LCase$(text)
    Find = LCase$(Find)
    Compare = vbBinaryCompare
    GoTo Begin
  End If
  
End Function

Public Function Replace(ByRef sText As String, ByRef sOld As String, ByRef sNew As String, Optional ByVal Start As Long = 1, Optional ByVal Count As Long = 2147483647, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As String
' by Jost Schwider, jost@schwider.de, 20001218
    
    If LenB(sOld) Then
 
        If Compare = vbBinaryCompare Then
            Replace2Bin Replace, sText, LCase$(sText), LCase$(sOld), sNew, Start, Count
        Else
            Replace2Bin Replace, sText, sText, sOld, sNew, Start, Count
        End If
    
    Else 'Suchsing ist leer:
        Replace = sText
    End If
  
End Function

Public Sub DoReplace(ByRef sText As String, ByRef sOld As String, ByRef sNew As String, Optional ByVal Start As Long = 1, Optional ByVal Count As Long = 2147483647, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare)
    
    If LenB(sOld) Then
        
        If Compare = vbBinaryCompare Then
            Replace3Bin sText, sText, sText, sOld, sNew, Start, Count
        Else
            Replace3Bin sText, sText, LCase$(sText), LCase$(sOld), sNew, Start, Count
        End If
        
    End If
    
End Sub

Private Sub Replace2Bin(ByRef Result As String, ByRef sText As String, ByRef Search As String, ByRef sOld As String, ByRef sNew As String, ByVal Start As Long, ByVal Count As Long)
' by Jost Schwider, jost@schwider.de, 20001218
Dim TextLen As Long
Dim OldLen As Long
Dim NewLen As Long
Dim ReadPos As Long
Dim WritePos As Long
Dim CopyLen As Long
Dim Buffer As String
Dim BufferLen As Long
Dim BufferPosNew As Long
Dim BufferPosNext As Long
  
    'Ersten Treffer bestimmen:
    If Start < 2 Then
        Start = InStrB(Search, sOld)
    Else
        Start = InStrB(Start + Start - 1, Search, sOld)
    End If
  
    If Start Then
  
        OldLen = LenB(sOld)
        NewLen = LenB(sNew)
    
        Select Case NewLen
    
            Case OldLen 'einfaches Überschreiben:
            
                Result = sText
            
                For Count = 1 To Count
                    MidB$(Result, Start) = sNew
                    Start = InStrB(Start + OldLen, Search, sOld)
                    If Start = 0 Then Exit Sub
                Next Count
                
                Exit Sub
    
            Case Is < OldLen 'Ergebnis wird kürzer:
        
                ' Initialize buffer
                TextLen = LenB(sText)
                
                If TextLen > BufferLen Then
                    Buffer = sText
                    BufferLen = TextLen
                End If
      
                'Ersetzen:
                ReadPos = 1
                WritePos = 1
            
                If NewLen Then
      
                    'Einzufügenden Text beachten:
                    For Count = 1 To Count
                
                        CopyLen = Start - ReadPos
                        
                        If CopyLen Then
                            BufferPosNew = WritePos + CopyLen
                            MidB$(Buffer, WritePos) = MidB$(sText, ReadPos, CopyLen)
                            MidB$(Buffer, BufferPosNew) = sNew
                            WritePos = BufferPosNew + NewLen
                        Else
                            MidB$(Buffer, WritePos) = sNew
                            WritePos = WritePos + NewLen
                        End If
                        
                        ReadPos = Start + OldLen
                        Start = InStrB(ReadPos, Search, sOld)
                        
                        If Start = 0 Then Exit For
                
                    Next Count
      
                Else
      
                    'Einzufügenden Text ignorieren (weil leer):
                    For Count = 1 To Count
            
                        CopyLen = Start - ReadPos
                        
                        If CopyLen Then
                            MidB$(Buffer, WritePos) = MidB$(sText, ReadPos, CopyLen)
                            WritePos = WritePos + CopyLen
                        End If
                      
                        ReadPos = Start + OldLen
                        Start = InStrB(ReadPos, Search, sOld)
                        
                        If Start = 0 Then Exit For
                        
                    Next Count
      
                End If
      
            'Ergebnis zusammenbauen:
            If ReadPos <= TextLen Then
                MidB$(Buffer, WritePos) = MidB$(sText, ReadPos)
                Result = LeftB$(Buffer, WritePos + LenB(sText) - ReadPos)
            Else
                Result = LeftB$(Buffer, WritePos - 1)
            End If
    
            Exit Sub
    
        Case Else 'Ergebnis wird länger:
    
            ' Initialize buffer
            TextLen = LenB(sText)
            BufferPosNew = TextLen + NewLen
            
            If BufferPosNew > BufferLen Then
                Buffer = Space$(BufferPosNew)
                BufferLen = LenB(Buffer)
            End If
      
            'Ersetzung:
            ReadPos = 1
            WritePos = 1
            
            For Count = 1 To Count
                
                CopyLen = Start - ReadPos
                
                If CopyLen Then
                    
                    'Positionen berechnen:
                    BufferPosNew = WritePos + CopyLen
                    BufferPosNext = BufferPosNew + NewLen
          
                    'Ggf. Buffer vergrößern:
                    If BufferPosNext > BufferLen Then
                        Buffer = Buffer & Space$(BufferPosNext)
                        BufferLen = LenB(Buffer)
                    End If
          
                    'String "patchen":
                    MidB$(Buffer, WritePos) = MidB$(sText, ReadPos, CopyLen)
                    MidB$(Buffer, BufferPosNew) = sNew
                    
                Else
                
                    'Position bestimmen:
                    BufferPosNext = WritePos + NewLen
          
                    'Ggf. Buffer vergrößern:
                    If BufferPosNext > BufferLen Then
                        Buffer = Buffer & Space$(BufferPosNext)
                        BufferLen = LenB(Buffer)
                    End If
          
                    'String "patchen":
                    MidB$(Buffer, WritePos) = sNew
                    
                End If
                
                WritePos = BufferPosNext
                ReadPos = Start + OldLen
                Start = InStrB(ReadPos, Search, sOld)
            
                If Start = 0 Then Exit For
            
            Next Count
      
            'Ergebnis zusammenbauen:
            If ReadPos <= TextLen Then
                
                BufferPosNext = WritePos + TextLen - ReadPos
                
                If BufferPosNext < BufferLen Then
                    MidB$(Buffer, WritePos) = MidB$(sText, ReadPos)
                    Result = LeftB$(Buffer, BufferPosNext)
                Else
                    Result = LeftB$(Buffer, WritePos - 1) & MidB$(sText, ReadPos)
                End If
                
            Else
                Result = LeftB$(Buffer, WritePos - 1)
            End If
            
            Exit Sub
    
        End Select
  
    Else 'Kein Treffer:
        Result = sText
    End If
  
End Sub

Private Sub Replace3Bin(ByRef Result As String, ByRef sText As String, ByRef Search As String, ByRef sOld As String, ByRef sNew As String, ByVal Start As Long, ByVal Count As Long)
' by Jost Schwider, jost@schwider.de, 20001218
Dim TextLen As Long
Dim OldLen As Long
Dim NewLen As Long
Dim ReadPos As Long
Dim WritePos As Long
Dim CopyLen As Long
Dim Buffer As String
Dim BufferLen As Long
Dim BufferPosNew As Long
Dim BufferPosNext As Long
  
    'Ersten Treffer bestimmen:
    If Start < 2 Then
        Start = InStrB(Search, sOld)
    Else
        Start = InStrB(Start + Start - 1, Search, sOld)
    End If
  
    If Start Then
  
        OldLen = LenB(sOld)
        NewLen = LenB(sNew)
    
        Select Case NewLen
    
            Case OldLen 'einfaches Überschreiben:
            
                For Count = 1 To Count
                    MidB$(Result, Start) = sNew
                    Start = InStrB(Start + OldLen, Search, sOld)
                    If Start = 0 Then Exit Sub
                Next Count
                
                Exit Sub
    
            Case Is < OldLen 'Ergebnis wird kürzer:
        
                ' Initialize buffer
                TextLen = LenB(sText)
                
                If TextLen > BufferLen Then
                    Buffer = sText
                    BufferLen = TextLen
                End If
      
                'Ersetzen:
                ReadPos = 1
                WritePos = 1
            
                If NewLen Then
      
                    'Einzufügenden Text beachten:
                    For Count = 1 To Count
                
                        CopyLen = Start - ReadPos
                        
                        If CopyLen Then
                            BufferPosNew = WritePos + CopyLen
                            MidB$(Buffer, WritePos) = MidB$(sText, ReadPos, CopyLen)
                            MidB$(Buffer, BufferPosNew) = sNew
                            WritePos = BufferPosNew + NewLen
                        Else
                            MidB$(Buffer, WritePos) = sNew
                            WritePos = WritePos + NewLen
                        End If
                        
                        ReadPos = Start + OldLen
                        Start = InStrB(ReadPos, Search, sOld)
                        
                        If Start = 0 Then Exit For
                
                    Next Count
      
                Else
      
                    'Einzufügenden Text ignorieren (weil leer):
                    For Count = 1 To Count
            
                        CopyLen = Start - ReadPos
                        
                        If CopyLen Then
                            MidB$(Buffer, WritePos) = MidB$(sText, ReadPos, CopyLen)
                            WritePos = WritePos + CopyLen
                        End If
                      
                        ReadPos = Start + OldLen
                        Start = InStrB(ReadPos, Search, sOld)
                        
                        If Start = 0 Then Exit For
                        
                    Next Count
      
                End If
      
            'Ergebnis zusammenbauen:
            If ReadPos <= TextLen Then
                MidB$(Buffer, WritePos) = MidB$(sText, ReadPos)
                Result = LeftB$(Buffer, WritePos + LenB(sText) - ReadPos)
            Else
                Result = LeftB$(Buffer, WritePos - 1)
            End If
    
            Exit Sub
    
        Case Else 'Ergebnis wird länger:
    
            ' Initialize buffer
            TextLen = LenB(sText)
            BufferPosNew = TextLen + NewLen
            
            If BufferPosNew > BufferLen Then
                Buffer = Space$(BufferPosNew)
                BufferLen = LenB(Buffer)
            End If
      
            'Ersetzung:
            ReadPos = 1
            WritePos = 1
            
            For Count = 1 To Count
                
                CopyLen = Start - ReadPos
                
                If CopyLen Then
                    
                    'Positionen berechnen:
                    BufferPosNew = WritePos + CopyLen
                    BufferPosNext = BufferPosNew + NewLen
          
                    'Ggf. Buffer vergrößern:
                    If BufferPosNext > BufferLen Then
                        Buffer = Buffer & Space$(BufferPosNext)
                        BufferLen = LenB(Buffer)
                    End If
          
                    'String "patchen":
                    MidB$(Buffer, WritePos) = MidB$(sText, ReadPos, CopyLen)
                    MidB$(Buffer, BufferPosNew) = sNew
                    
                Else
                
                    'Position bestimmen:
                    BufferPosNext = WritePos + NewLen
          
                    'Ggf. Buffer vergrößern:
                    If BufferPosNext > BufferLen Then
                        Buffer = Buffer & Space$(BufferPosNext)
                        BufferLen = LenB(Buffer)
                    End If
          
                    'String "patchen":
                    MidB$(Buffer, WritePos) = sNew
                    
                End If
                
                WritePos = BufferPosNext
                ReadPos = Start + OldLen
                Start = InStrB(ReadPos, Search, sOld)
            
                If Start = 0 Then Exit For
            
            Next Count
      
            'Ergebnis zusammenbauen:
            If ReadPos <= TextLen Then
                
                BufferPosNext = WritePos + TextLen - ReadPos
                
                If BufferPosNext < BufferLen Then
                    MidB$(Buffer, WritePos) = MidB$(sText, ReadPos)
                    Result = LeftB$(Buffer, BufferPosNext)
                Else
                    Result = LeftB$(Buffer, WritePos - 1) & MidB$(sText, ReadPos)
                End If
                
            Else
                Result = LeftB$(Buffer, WritePos - 1)
            End If
            
            Exit Sub
    
        End Select
  
    End If
  
End Sub

Public Sub SplitB(sExpression As String, sSplitArray() As String, Optional sDelimiter As String = " ", Optional lKeyCount As Long, Optional BUFFERDIM As Long = 10)
' by Donald, donald@xbeat.net, 20000916
' modified by Keith, kmatzen@ispchannel.com, 20000923
Dim lCntSplits    As Long
Dim lCntStart     As Long
Dim lUBound       As Long
Dim lPosStart     As Long
Dim lPosFound     As Long
Dim lLenDelimiter As Long
Dim lStrLen       As Long
   
    lLenDelimiter = Len(sDelimiter)
    lPosFound = InStr(sExpression, sDelimiter)
   
    If lLenDelimiter = 0 Or lPosFound = 0 Then
   
        ' No delimiters - return sExpression in single-element array
        ReDim Preserve sSplitArray(0)
        sSplitArray(0) = sExpression
        lKeyCount = 1
        Exit Sub
      
    End If
    
    lPosStart = 1
    lUBound = -1
    
    Do
       
        lCntStart = lUBound + 1
        lUBound = lUBound + BUFFERDIM
        ReDim Preserve sSplitArray(lUBound)
       
        For lCntSplits = lCntStart To lUBound
            
            If lPosFound Then
                ' Delimiter found
                lStrLen = lPosFound - lPosStart
                sSplitArray(lCntSplits) = Mid$(sExpression, lPosStart, lStrLen)
                lPosStart = lPosFound + lLenDelimiter
                lPosFound = InStr(lPosStart, sExpression, sDelimiter)
            Else
                ' No more delimiters
                sSplitArray(lCntSplits) = Mid$(sExpression, lPosStart)
                ReDim Preserve sSplitArray(lCntSplits)
                lKeyCount = lCntSplits + 1
                Exit Sub
            End If
          
       Next lCntSplits
       
    Loop
  
End Sub

Public Sub SafeClipboardSet(sString As String)
    
    On Error Resume Next
    Sleep 5: Clipboard.Clear
    Sleep 5: Clipboard.SetText sString

End Sub

Public Sub Show2(frmToShow As Form, frmParent As Form, TopMost As Boolean, Optional Modal As FormShowConstants = vbModeless)
    
    On Error GoTo Hell

    If TopMost = True Then
        SetTopmostWindow frmParent.hWnd, False
    End If
    
    frmToShow.Show Modal, frmParent
    
    If TopMost = True Then
        SetTopmostWindow frmParent.hWnd, True, True
    End If
    
Hell:
End Sub

Public Function WriteToPrevInstance(EditText As String, Title As String) As Boolean
'This function searches all top level windows to find one with a matching Window Title
'Then all of the window's child windows are searched for a "ThunderTextbox" class window
'The text is then written to the Edit window

    'save the search string to a module level variable
    sWindowTitle = Title
    
    'Save the new text to a module variable
    sNewText = EditText
    
    'start searching...
    EnumWindows AddressOf EnumWindowProc, &H0
    
    WriteToPrevInstance = Success
    
End Function

Private Function EnumWindowProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    
    'eliminate windows that are not top-level.
    If GetParent(hWnd) = 0& Then
        
        'See if we have found our window
        If InStrB(GetWindowTitle(hWnd), sWindowTitle) <> 0 Then
         
            'Excellent, now search for a ThunderTextbox window in this app
            EnumChildWindows hWnd, AddressOf EnumChildProc, &H0
         
            'stop looking
            EnumWindowProc = False
            Exit Function
            
        End If
       
    End If
    
    'To continue enumeration, return True
    'To stop enumeration return False (0).
    'When 1 is returned, enumeration continues
    'until there are no more windows left.
    EnumWindowProc = True
    
End Function

Private Function GetWindowTitle(ByVal hWnd As Long) As String
Dim lSize As Long
Dim sTitle As String
    
    'get the size of the string required
    'to hold the window title
    lSize = GetWindowTextLength(hWnd)
    
    'if the return is 0, there is no title
    If lSize > 0 Then
        sTitle = Space$(lSize + 1)
        GetWindowText hWnd, sTitle, lSize + 1
    End If
    
    GetWindowTitle = sTitle

End Function

Public Function HasPrinters() As Boolean
Dim cbRequired As Long
Dim lBuffer() As Long
Dim lEntries As Long
    
    On Error GoTo Hell
    
    EnumPrinters PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, 1, 0, 0, cbRequired, lEntries
    
    ReDim lBuffer(cbRequired) As Long
    EnumPrinters PRINTER_ENUM_CONNECTIONS Or PRINTER_ENUM_LOCAL, vbNullString, 1, lBuffer(0), cbRequired, cbRequired, lEntries
    
    If lEntries > 0 Then
        HasPrinters = True
    End If

Hell:
End Function

Private Function EnumChildProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
Const WS_VISIBLE = &H10000000
Dim sClass As String
Dim lStyle As Long
    
    sClass = Space$(20) 'enough for ThunderRT6TextBox
    GetClassName hWnd, sClass, 20
    lStyle = GetWindowLong(hWnd, GWL_STYLE) And WS_VISIBLE
    
    'Look for an edit window
    If InStrB(sClass, "TextBox") And (lStyle <> WS_VISIBLE) Then
       
       'set the new text
       SendMessage hWnd, WM_SETTEXT, 0&, ByVal sNewText
       Success = True
       
       'Done looking
       EnumChildProc = 0
       Exit Function
       
    Else
        'keep looking
        EnumChildProc = 1
    End If
   
End Function

Public Sub BlastText(txtTextBox As TextBox, sFileName As String)
Dim iFileNum As Integer
Dim sTemp As String
Dim cString As cStringBuilder

    iFileNum = FreeFile

    Open sFileName For Input As #iFileNum
        On Error GoTo AlternateWay
        sTemp = Input$(LOF(iFileNum), iFileNum)
    Close #iFileNum

    GoTo Finalize

AlternateWay:
    
    Close #iFileNum

    iFileNum = FreeFile
    Set cString = New cStringBuilder
    
    Open sFileName For Input As #iFileNum
        Do While Not EOF(iFileNum)
            Line Input #iFileNum, sTemp
            cString.Append sTemp & vbNewLine
        Loop
    Close #iFileNum
    
    sTemp = cString.ToString
    cString.Clear
    
    If Len(sTemp) > FileLength(sFileName) Then
        sTemp = Left$(sTemp, FileLength(sFileName))
    End If
    
    GoTo Finalize

Finalize:

    SendMessageW txtTextBox.hWnd, WM_SETTEXT, 0&, ByVal StrPtr(sTemp)

End Sub
