Attribute VB_Name = "modUnlockFile"
Option Explicit

'// This code was originally downloaded from: http://blog.csdn.net/chenhui530/archive/2007/12/13/1932917.aspx
'// All credits go to the original author.

'// Unnecessary code has been removed and some code modified.

Private Enum SYSTEM_INFORMATION_CLASS
    SystemBasicInformation
    SystemProcessorInformation
    SystemPerformanceInformation
    SystemTimeOfDayInformation
    SystemPathInformation
    SystemProcessInformation
    SystemCallCountInformation
    SystemDeviceInformation
    SystemProcessorPerformanceInformation
    SystemFlagsInformation
    SystemCallTimeInformation
    SystemModuleInformation
    SystemLocksInformation
    SystemStackTraceInformation
    SystemPagedPoolInformation
    SystemNonPagedPoolInformation
    SystemHandleInformation
    SystemObjectInformation
    SystemPageFileInformation
    SystemVdmInstemulInformation
    SystemVdmBopInformation
    SystemFileCacheInformation
    SystemPoolTagInformation
    SystemInterruptInformation
    SystemDpcBehaviorInformation
    SystemFullMemoryInformation
    SystemLoadGdiDriverInformation
    SystemUnloadGdiDriverInformation
    SystemTimeAdjustmentInformation
    SystemSummaryMemoryInformation
    SystemMirrorMemoryInformation
    SystemPerformanceTraceInformation
    SystemObsolete0
    SystemExceptionInformation
    SystemCrashDumpStateInformation
    SystemKernelDebuggerInformation
    SystemContextSwitchInformation
    SystemRegistryQuotaInformation
    SystemExtendServiceTableInformation
    SystemPrioritySeperation
    SystemVerifierAddDriverInformation
    SystemVerifierRemoveDriverInformation
    SystemProcessorIdleInformation
    SystemLegacyDriverInformation
    SystemCurrentTimeZoneInformation
    SystemLookasideInformation
    SystemTimeSlipNotification
    SystemSessionCreate
    SystemSessionDetach
    SystemSessionInformation
    SystemRangeStartInformation
    SystemVerifierInformation
    SystemVerifierThunkExtend
    SystemSessionProcessInformation
    SystemLoadGdiDriverInSystemSpace
    SystemNumaProcessorMap
    SystemPrefetcherInformation
    SystemExtendedProcessInformation
    SystemRecommendedSharedDataAlignment
    SystemComPlusPackage
    SystemNumaAvailableMemory
    SystemProcessorPowerInformation
    SystemEmulationBasicInformation
    SystemEmulationProcessorInformation
    SystemExtendedHandleInformation
    SystemLostDelayedWriteInformation
    SystemBigPoolInformation
    SystemSessionPoolTagInformation
    SystemSessionMappedViewInformation
    SystemHotpatchInformation
    SystemObjectSecurityMode
    SystemWatchdogTimerHandler
    SystemWatchdogTimerInformation
    SystemLogicalProcessorInformation
    SystemWow64SharedInformation
    SystemRegisterFirmwareTableInformationHandler
    SystemFirmwareTableInformation
    SystemModuleInformationEx
    SystemVerifierTriageInformation
    SystemSuperfetchInformation
    SystemMemoryListInformation
    SystemFileCacheInformationEx
    MaxSystemInfoClass  '// MaxSystemInfoClass should always be the last enum
End Enum

Private Type SYSTEM_HANDLE
    UniqueProcessId As Integer
    CreatorBackTraceIndex As Integer
    ObjectTypeIndex As Byte
    HandleAttributes As Byte
    HandleValue As Integer
    pObject As Long
    GrantedAccess As Long
End Type

Private Type SYSTEM_HANDLE_INFORMATION
    uCount As Long
    aSH() As SYSTEM_HANDLE
End Type

Private Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As Long
    Attributes As Long
    SecurityDescriptor As Long
    SecurityQualityOfService As Long
End Type

Private Type CLIENT_ID
    UniqueProcess As Long
    UniqueThread  As Long
End Type

Private Enum OBJECT_INFORMATION_CLASS
    ObjectBasicInformation = 0
    ObjectNameInformation
    ObjectTypeInformation
    ObjectAllTypesInformation
    ObjectHandleInformation
End Enum

Private Type UNICODE_STRING
    uLength As Integer
    uMaximumLength As Integer
    pBuffer(3) As Byte
End Type

Private Const STATUS_INFO_LENGTH_MISMATCH = &HC0000004
Private Const DUPLICATE_CLOSE_SOURCE = &H1
Private Const DUPLICATE_SAME_ACCESS = &H2
Private Const DUPLICATE_SAME_ATTRIBUTES = &H4
Private Const PROCESS_DUP_HANDLE As Long = &H40&

Private Const STATUS_INFO_LEN_MISMATCH = &HC0000004
Private Const HEAP_ZERO_MEMORY = &H8

Private Const WM_CLOSE As Long = &H10&

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As Any, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function NtQuerySystemInformation Lib "NTDLL.DLL" (ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, ByVal pSystemInformation As Long, ByVal SystemInformationLength As Long, ByRef ReturnLength As Long) As Long
Private Declare Function NtQueryObject Lib "NTDLL.DLL" (ByVal ObjectHandle As Long, ByVal ObjectInformationClass As OBJECT_INFORMATION_CLASS, ByVal ObjectInformation As Long, ByVal ObjectInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function NtDuplicateObject Lib "NTDLL.DLL" (ByVal SourceProcessHandle As Long, ByVal SourceHandle As Long, ByVal TargetProcessHandle As Long, ByRef TargetHandle As Long, ByVal DesiredAccess As Long, ByVal HandleAttributes As Long, ByVal Options As Long) As Long
Private Declare Function NtOpenProcess Lib "NTDLL.DLL" (ByRef ProcessHandle As Long, ByVal AccessMask As Long, ByRef ObjectAttributes As OBJECT_ATTRIBUTES, ByRef ClientID As CLIENT_ID) As Long
Private Declare Function NtClose Lib "NTDLL.DLL" (ByVal ObjectHandle As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Function GetFileFullPath(ByVal hFile As Long) As String
Dim hHeap As Long
Dim dwSize As Long
Dim objName As UNICODE_STRING
Dim pName As Long
Dim ntStatus As Long
Dim i As Long
Dim sDrives As String
Dim sArray() As String
Dim dwDriversSize As Long
Dim sDrive As String
Dim sTmp As String
Dim sTemp As String
    
    hHeap = GetProcessHeap
    pName = HeapAlloc(hHeap, HEAP_ZERO_MEMORY, &H1000&)
    ntStatus = NtQueryObject(hFile, ObjectNameInformation, pName, &H1000&, dwSize)
    
    If (ntStatus >= 0) Then
        i = 1
        Do While (ntStatus = STATUS_INFO_LEN_MISMATCH)
            pName = HeapReAlloc(hHeap, HEAP_ZERO_MEMORY, pName, &H1000& * i)
            ntStatus = NtQueryObject(hFile, ObjectNameInformation, pName, &H1000&, ByVal 0&)
            i = i + 1
        Loop
    End If
    
    HeapFree hHeap, 0, pName
    sTemp = String$(512, vbNullChar)
    lstrcpyW sTemp, pName + Len(objName)
    sTemp = StrConv(sTemp, vbFromUnicode)
    sTemp = Left$(sTemp, InStr(sTemp, vbNullChar) - 1)
    sDrives = Space$(512)
    dwDriversSize = GetLogicalDriveStrings(512, sDrives)
    
    If dwDriversSize Then
        sArray = Split(sDrives, vbNullChar)
        For i = LBound(sArray) To UBound(sArray)
            If LenB(sArray(i)) <> 0 Then
                sDrive = Left$(sArray(i), 2)
                sTmp = String$(260, vbNullChar)
                Call QueryDosDevice(sDrive, sTmp, 256)
                sTmp = Left$(sTmp, InStr(sTmp, vbNullChar) - 1)
                If InStrB(LCase$(sTemp), LCase$(sTmp)) = 1 Then
                    GetFileFullPath = sDrive & Mid$(sTemp, Len(sTmp) + 1, Len(sTemp) - Len(sTmp))
                    Exit Function
                End If
            End If
        Next i
    End If

End Function

Public Function UnLockFile(sFileName As String) As Boolean
Dim ntStatus As Long
Dim objCid As CLIENT_ID
Dim objOa As OBJECT_ATTRIBUTES
Dim lHandles As Long
Dim i As Long
Dim objInfo As SYSTEM_HANDLE_INFORMATION
Dim lType As Long
Dim hProcessToDup As Long
Dim sFile As String
Dim hFileHandle As Long
Dim hFile As Long
Dim sTmp As String

    hFile = CreateFile("NUL", &H80000000, 0, ByVal 0&, 3, 0, 0)
    
    If hFile = -1 Then
        UnLockFile = False
        Exit Function
    End If
    
    sFile = sFileName
    objOa.Length = Len(objOa)
    ntStatus = 0
    
    Dim bBuf() As Byte
    Dim nSize As Long
    
    nSize = 1
    
    Do
        ReDim bBuf(nSize)
        ntStatus = NtQuerySystemInformation(SystemHandleInformation, VarPtr(bBuf(0)), nSize, 0&)
        If (Not ntStatus >= 0) Then
            If (ntStatus <> STATUS_INFO_LENGTH_MISMATCH) Then
                Erase bBuf
                Exit Function
            End If
        Else
            Exit Do
        End If
        nSize = nSize * 2
        ReDim bBuf(nSize)
    Loop
    
    lHandles = 0
    RtlMoveMemory objInfo.uCount, bBuf(0), 4
    lHandles = objInfo.uCount
    
    ReDim objInfo.aSH(lHandles - 1)
    RtlMoveMemory objInfo.aSH(0), bBuf(4), Len(objInfo.aSH(0)) * lHandles
    
    For i = 0 To lHandles - 1
        If objInfo.aSH(i).HandleValue = hFile And objInfo.aSH(i).UniqueProcessId = GetCurrentProcessId Then
            lType = objInfo.aSH(i).ObjectTypeIndex
            Exit For
        End If
    Next i
    
    NtClose hFile
    UnLockFile = True
    
    For i = 0 To lHandles - 1
        If objInfo.aSH(i).ObjectTypeIndex = lType Then
            objCid.UniqueProcess = objInfo.aSH(i).UniqueProcessId
            ntStatus = NtOpenProcess(hProcessToDup, PROCESS_DUP_HANDLE, objOa, objCid)
            If hProcessToDup <> 0 Then
                ntStatus = NtDuplicateObject(hProcessToDup, objInfo.aSH(i).HandleValue, GetCurrentProcess, hFileHandle, 0, 0, DUPLICATE_SAME_ATTRIBUTES)
                If (ntStatus >= 0) Then
                    ntStatus = MyGetFileType(hFileHandle)
                    If ntStatus Then
                        sTmp = GetFileFullPath(hFileHandle)
                    Else
                        sTmp = vbNullString
                    End If
                    NtClose hFileHandle
                    If InStrB(LCase$(sTmp), LCase$(sFile)) Then
                        If Not CloseRemoteHandle(objInfo.aSH(i).UniqueProcessId, objInfo.aSH(i).HandleValue) Then
                            UnLockFile = False
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
End Function

Private Function GetProcessCommandLine() As String
Dim objOa As OBJECT_ATTRIBUTES
Dim hKernel As Long
Dim sName As String
Dim hProcess As Long
Dim dwAddr As Long
Dim dwRead As Long

    objOa.Length = Len(objOa)
    
    If hProcess = 0 Then
        GetProcessCommandLine = vbNullString
        Exit Function
    End If
    
    hKernel = GetModuleHandle("kernel32")
    dwAddr = GetProcAddress(hKernel, "GetCommandLineA")
    RtlMoveMemory dwAddr, ByVal dwAddr + 1, 4
    
    If ReadProcessMemory(hProcess, ByVal dwAddr, dwAddr, 4, dwRead) Then
        sName = String$(260, vbNullChar)
        If ReadProcessMemory(hProcess, ByVal dwAddr, ByVal sName, 260, dwRead) Then
            sName = Left$(sName, InStr(sName, vbNullChar) - 1)
            NtClose hProcess
            GetProcessCommandLine = sName
            Exit Function
        End If
    End If
    
    NtClose hProcess
    
End Function

Private Function CloseRemoteHandle(ByVal dwProcessId As Long, ByVal hHandle As Long) As Boolean
Dim hRemProcess As Long
Dim lResult As Long
Dim hMyHandle As Long
Dim objCid As CLIENT_ID
Dim objOa As OBJECT_ATTRIBUTES
Dim ntStatus As Long
Dim sProcessName As String
Dim hProcess As Long
    
    objCid.UniqueProcess = dwProcessId
    objOa.Length = Len(objOa)
    ntStatus = NtOpenProcess(hRemProcess, PROCESS_DUP_HANDLE, objOa, objCid)
    
    If hRemProcess Then
        ntStatus = NtDuplicateObject(hRemProcess, hHandle, GetCurrentProcess, hMyHandle, 0, 0, DUPLICATE_CLOSE_SOURCE Or DUPLICATE_SAME_ACCESS)
        If (ntStatus >= 0) Then
            lResult = NtClose(hMyHandle)
            If lResult >= 0 Then
                sProcessName = GetProcessCommandLine()
                If InStrB(LCase$(sProcessName), "explorer.exe") = 0 And dwProcessId <> GetCurrentProcessId Then
                    objCid.UniqueProcess = dwProcessId
                    ntStatus = NtOpenProcess(hProcess, 1, objOa, objCid)
                    If hProcess <> 0 Then PostMessage hProcess, WM_CLOSE, 0&, 0&
                End If
            End If
        End If
        Call NtClose(hRemProcess)
    End If
    
    CloseRemoteHandle = lResult >= 0
    
End Function

Private Function MyGetFileType(ByVal hFile As Long) As Long
Dim hRemProcess As Long, hThread As Long, lResult As Long, pfnThreadRtn As Long, hKernel As Long
Dim dwEax As Long, dwTimeOut As Long
    
    hRemProcess = GetCurrentProcess
    hKernel = GetModuleHandle("kernel32")
    
    If hKernel = 0 Then
        MyGetFileType = 0
        Exit Function
    End If
    
    pfnThreadRtn = GetProcAddress(hKernel, "GetFileType")
    
    If pfnThreadRtn = 0 Then
        FreeLibrary hKernel
        MyGetFileType = 0
        Exit Function
    End If
    
    hThread = CreateRemoteThread(hRemProcess, ByVal 0&, 0&, ByVal pfnThreadRtn, ByVal hFile, 0, ByVal 0&)
    dwEax = WaitForSingleObject(hThread, 100)
    
    If dwEax = &H102& Then
        Call GetExitCodeThread(hThread, dwTimeOut)
        Call TerminateThread(hThread, dwTimeOut)
        NtClose hThread
        MyGetFileType = 0
        Exit Function
    End If
    
    If hThread = 0 Then
        FreeLibrary hKernel
        MyGetFileType = 0
        Exit Function
    End If
    
    GetExitCodeThread hThread, lResult
    MyGetFileType = lResult
    NtClose hThread
    NtClose hRemProcess
    FreeLibrary hKernel
    
End Function
