VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ****************************************************
'
'   cStringBuilder
'   By VolteFace
'
'   Date Created: 3/21/2004
'
'   This class was created to provide more or less the
'   same functionality as the System.Text.StringBuider
'   class available in the .NET framework. It makes use
'   of direct memory allocation and manipulation, so is
'   much faster than traditional VB string concatenation.
'
' ****************************************************

' API Declares
Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapReAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub RtlZeroMemory Lib "kernel32.dll" (Destination As Any, ByVal Length As Long)
Private Declare Function SysAllocString Lib "oleaut32.dll" (pOlechar As Long) As String
Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long

' Contants
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const CHUNK_SIZE = 1048576

' Member Variables
Private m_pMemoryPtr As Long
Private m_lAllocSize As Long
Private m_lLength As Long

' Initializes the class and allocates the
' initial string buffer.
Private Sub Class_Initialize()
    ' Allocate 1MB by default
    Allocate CHUNK_SIZE
End Sub

' Allocates a specified amount of memory
' for the string buffer.
Private Sub Allocate(ByVal SIZE As Long)
    
    ' If no memory is allocated yet, allocate some from the heap - otherwise
    ' reallocate (resize) the block that has already been allocated
    If m_pMemoryPtr = 0& Then
        m_pMemoryPtr = HeapAlloc(GetProcessHeap, HEAP_ZERO_MEMORY, SIZE)
    Else
        m_pMemoryPtr = HeapReAlloc(GetProcessHeap, 0&, ByVal m_pMemoryPtr, SIZE)
    End If
    
    m_lAllocSize = SIZE
    
End Sub

' Gets the internally stored string
' data as a VB String variable.
Public Property Get ToString() As String
    ' Create a buffer that is the size of the stored string and
    ' copy the string contents at the stored memory pointer into the buffer
    ToString = StrConv(SysAllocString(ByVal m_pMemoryPtr), vbFromUnicode)
End Property

' Sets the internally stored string
Public Property Let TheString(ByRef Value As String)
    Clear
    Append Value
End Property

' Removes all string data from the
' initial string buffer
Public Sub Reset()
    ' Clean out the string buffer
    RtlZeroMemory ByVal m_pMemoryPtr, m_lLength
    m_lLength = 0&
End Sub

' Removes all string data from the
' initial string buffer, and resizes
' the buffer down to the initial 1MB.
Public Sub Clear()
    Reset
    Allocate CHUNK_SIZE
End Sub

' Adds a specified string on to the
' end of the string stored in the buffer.
Public Sub Append(ByRef str As String)
Dim pTo As Long

    ' If we are going to need more memory (if the final size of the append is going to be
    ' greater than the currently allocated size), we need to find out how much more we
    ' need (in increments of CHUNK_SIZE, default 1MB) and allocate it
    If m_lLength + LenB(str) > m_lAllocSize Then
        Allocate m_lAllocSize + (CHUNK_SIZE * Int((m_lLength + LenB(str)) \ m_lAllocSize))
    End If
    
    ' Put the specified string at the end of the string buffer
    pTo = m_pMemoryPtr + m_lLength
    RtlMoveMemory ByVal pTo, ByVal StrPtr(str), LenB(str)
    
    m_lLength = m_lLength + LenB(str)
    
End Sub

' Inserts a specified string into the
' stored string at a specific index.
Public Sub Insert(ByVal Index As Long, ByRef str As String)
Dim pFrom As Long
Dim pTo As Long

    ' If we are going to need more memory (if the final size of the insert is going to be
    ' greater than the currently allocated size), we need to find out how much more we
    ' need (in increments of CHUNK_SIZE, default 1MB) and allocate it
    If m_lLength + LenB(str) > m_lAllocSize Then
        Allocate m_lAllocSize + (CHUNK_SIZE * Int((m_lLength + LenB(str)) \ m_lAllocSize))
    End If
    
    ' Copy the entire stored string, from 'index' to the end and move it over to the
    ' right to accomodate for the new string to be inserted, and then put the specified
    ' string in the correct position
    If (Index >= 0) And (Index <= (m_lLength \ 2&)) Then
        
        pFrom = m_pMemoryPtr + (Index * 2&)
        pTo = m_pMemoryPtr + (Index * 2&) + LenB(str)
    
        RtlMoveMemory ByVal pTo, ByVal pFrom, m_lLength - (Index * 2&)
        RtlMoveMemory ByVal pFrom, ByVal StrPtr(str), LenB(str)
        
        m_lLength = m_lLength + LenB(str)
        
    End If
    
End Sub

Public Function Find(ByRef sToFind As String, Optional Start As Long = -1&) As Long

    If LenB(sToFind) <= m_lLength Then
        If Start = -1 Then
            Find = (InStr(SysAllocString(ByVal m_pMemoryPtr), StrConv(sToFind, vbUnicode)) + 1&) \ 2&
        Else
            If Start <= Length Then
                Find = (InStr((Start * 2&) - 1&, SysAllocString(ByVal m_pMemoryPtr), StrConv(sToFind, vbUnicode)) + 1&) \ 2&
            End If
        End If
    End If
    
End Function

'' Inserts a string into the middle
'' of the stored string, wiping out
'' the characters at that position.
'Public Sub Overwrite(ByVal Index As Long, ByRef str As String)
'Dim pFrom As Long
'Dim pTo As Long
'
'    ' If we are going to need more memory (if the inserted string goes over
'    ' the length of the current string, and ends up being longer than the allocated
'    ' memory block, we need to calculate how much we need (in increments of CHUNK_SIZE,
'    ' default 1MB) and allocate it
'    If Index + LenB(str) > m_lLength Then
'        If m_lLength + Index + LenB(str) > m_lAllocSize Then
'            Allocate m_lAllocSize + (CHUNK_SIZE * Int((m_lLength + LenB(str)) \ m_lAllocSize))
'        End If
'    End If
'
'    ' Copy the specified string into the stored string
'    If (Index >= 0) And (Index <= (m_lLength \ 2)) Then
'
'        pFrom = m_pMemoryPtr + (Index * 2)
'
'        RtlMoveMemory ByVal pFrom, ByVal StrPtr(str), LenB(str)
'
'        ' If the string got longer (the inserted string hung over the end of the
'        ' old string) we need to calculate how much bigger it got
'        If (Index * 2) + LenB(str) > m_lLength Then
'            m_lLength = m_lLength + (((Index * 2) + LenB(str)) - m_lLength)
'        End If
'
'    End If
'
'End Sub

' Removes text from the middle of
' the stored string.
Public Sub Remove(ByVal Index As Long, ByVal Length As Long)
Dim pFrom As Long
Dim pTo As Long
Dim pEnd As Long
    
    ' Copy the entire stored string, from 'index' to the end and move it over to the
    ' left to overright the desired characters, and then excess characters at the end
    ' of the string
    If (Length >= 0&) And (Index < (m_lLength \ 2&)) And (Index >= 0&) Then
        
        If (Length + Index > (m_lLength \ 2&)) Or (Length = 0&) Then
            Length = (m_lLength \ 2&) - Index
        End If
    
        pTo = m_pMemoryPtr + (Index * 2&)
        pFrom = m_pMemoryPtr + ((Index + Length) * 2&)
        pEnd = m_pMemoryPtr + m_lLength - (Length * 2&)
        
        RtlMoveMemory ByVal pTo, ByVal pFrom, m_lLength - ((Index + Length) * 2&)
        RtlZeroMemory ByVal pEnd, Length * 2&
        
        m_lLength = m_lLength - (Length * 2&)
        
    End If
    
End Sub

' Returns the length of the string
Public Property Get Length() As Long
    ' Since the string is stored as unicode, every character is 2 bytes
    Length = m_lLength \ 2&
End Property

Public Property Get Pointer() As Long
    Pointer = m_pMemoryPtr
End Property

' Deallocates all allocated memory.
Private Sub Class_Terminate()
    ' If we have memory allocated, free it
    If m_pMemoryPtr <> 0& Then
        HeapFree GetProcessHeap, 0&, ByVal m_pMemoryPtr
    End If
End Sub
