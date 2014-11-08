Attribute VB_Name = "modScripting"
Option Explicit

Private Type tRubiParam
    SIZE As Byte
    Description As String
End Type

Private Type tRubiCommand
    NeededBytes As Byte
    ParamCount As Byte
    Keyword As String
    Description As String
End Type

Private Type tLevelScript
    Offset As Long
    Variable As Integer
    Value As Integer
    Pointer As Long
    Variable2 As Integer
End Type

Private Type tJunk
    Offset As Long
    Length As Long
End Type

Public RubiCommands() As tRubiCommand
Public RubiParams() As tRubiParam
Public RubiLookup() As Define

Public AutoBank As Boolean
Public NoLog As Boolean
Public IsFireRed As Boolean
Public Japanese As Boolean
Public IsLevelScript As Boolean
Public IsDebugging As Boolean
Public IsLookupReady As Boolean
Private TriedUnlocking As Boolean
Private MissingDynamic As Boolean
Private MissingDefine As Boolean

Private bFreeSpace As Byte
Public sTempPath As String
Public sGameCode As String * 3
Public MaxCommand As Byte
Private Const MaxKeywords = 255

Public iComments As Integer
Public iRefactoring As Integer
Public iDecompileMode As Integer
Public sCommentChar As String
Public sRefactorDynamic As String

Private Const sHexPrefix As String = " 0x"
Private Const ascAt As Integer = 64 '"@"
Public Const sTempLog As String = "~tmp.log"
Public Const sTempFile As String = "~tmp.rbc"

Private Enum iDecompileModes
    Strict = 0
    Normal = 1
    Enhanced = 2
End Enum

Private Enum Errors
    InvalidProcCall = 5
    Overflow = 6
    IndexOutRange = 9
    TypeMismatch = 13
    FileNotFound = 53
    BadRecordNumber = 63
    FileAccessErr = 75
    ObjectNotSet = 91
    DuplicateKey = 457
End Enum

Private Type ByteString
    sText As String
    bArray() As Byte
End Type

Private Defines As Collection
Private Headers As Collection
Private Aliases As Collection

Private Snips As Collection
Private Strings As Collection
Private Brailles As Collection
Private Moves As Collection
Private Marts As Collection

Private lDynamicCount As Long
Private DynamicOffsets2() As Define
Private JunkData() As tJunk

Private sMoveLabels() As String
Private MovesReady As Boolean

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public Sub InitCollections()
    Set Defines = New Collection
    Set Headers = New Collection
    Set Aliases = New Collection
    Set Snips = New Collection
    Set Strings = New Collection
    Set Moves = New Collection
    Set Marts = New Collection
    Set Brailles = New Collection
End Sub

Public Sub EraseCol(cCol As Collection)
    Set cCol = Nothing
    Set cCol = New Collection
End Sub

'---------------------------------------------------------------------------------------
' Procedure : CollectionKeyExists
' DateTime  : 11-5-2007
' Author    : Flyguy
'---------------------------------------------------------------------------------------
'Private Function ColKeyExists(ByVal sKey As String, cCol As Collection) As Boolean
'Dim bDummy As Boolean
'Dim lErr As Long
'
'    If Not cCol Is Nothing Then
'
'        Err.Clear
'        On Error Resume Next
'
'        bDummy = IsObject(cCol(sKey))
'        lErr = Err.Number
'
'        If lErr = 0 Then
'            ColKeyExists = True
'        ElseIf lErr = InvalidProcCall Then
'            ColKeyExists = False
'        Else
'            On Error GoTo 0
'            Err.Raise lErr
'        End If
'
'    'Else
'    '    Err.Raise ObjectNotSet
'    End If
'
'End Function

'---------------------------------------------------------------------------------------
' Procedure : ItemKey
' DateTime  : 11-5-2007
' Author    : LaVolpe
' Purpose   : Get collection key by index
'---------------------------------------------------------------------------------------
Private Function ColItemKey(ByVal Index As Long, Coll As Collection) As String
Dim i     As Long
Dim Ptr   As Long
Dim sKey  As String

    If Not Coll Is Nothing Then
        
        Select Case Index
            
            Case Is <= Coll.Count \ 2 'walk items upwards from first one
              
              RtlMoveMemory Ptr, ByVal ObjPtr(Coll) + 24, 4 'first Ptr
              
              For i = 2 To Index
                  RtlMoveMemory Ptr, ByVal Ptr + 24, 4 'next Ptr
              Next i
              
            Case Is > Coll.Count \ 2 'walk items downwards from last one

              RtlMoveMemory Ptr, ByVal ObjPtr(Coll) + 28, 4 'last Ptr
              
              For i = Coll.Count - 1 To Index Step -1
                  RtlMoveMemory Ptr, ByVal Ptr + 20, 4 'prev Ptr
              Next i
              
            Case Is < 1, Is > Coll.Count 'oops!
                Err.Raise IndexOutRange
              
        End Select
        
        i = StrPtr(sKey) 'save StrPtr
        RtlMoveMemory ByVal VarPtr(sKey), ByVal Ptr + 16, 4 'Replace StrPtr by that from collection sKey (which is null if there ain't no sKey)
        ColItemKey = sKey 'now copy it to Function value()
        RtlMoveMemory ByVal VarPtr(sKey), i, 4 'and finally restore original StrPtr
        
    'Else
    '    Err.Raise ObjectNotSet 'No object
    End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : ItemIndex
' DateTime  : 11-5-2007
' Author    : LaVolpe
' Purpose   : Get collection index by key
'---------------------------------------------------------------------------------------
'Public Function ColItemIndex(ByVal Key As String, Coll As Collection, Optional ByVal compare As VbCompareMethod = vbTextCompare) As Long
'Dim Ptr   As Long
'Dim sKey  As String
'Dim aKey  As Long
'
'    If Not Coll Is Nothing Then
'        If Coll.Count Then
'            aKey = StrPtr(sKey)                         'save StrPtr
'            RtlMoveMemory Ptr, ByVal ObjPtr(Coll) + 24, 4  'first Ptr
'            ColItemIndex = 1                            'walk items upwards From First
'            Do
'                RtlMoveMemory ByVal VarPtr(sKey), ByVal Ptr + 16, 4
'                If StrComp(Key, sKey, compare) = 0 Then 'equal
'                    Exit Do 'found
'                End If
'                ColItemIndex = ColItemIndex + 1  'next Index
'                RtlMoveMemory Ptr, ByVal Ptr + 24, 4              'next Ptr
'            Loop Until Ptr = 0                                 'end of chain
'            RtlMoveMemory ByVal VarPtr(sKey), aKey, 4             'restore original StrPtr
'        End If
'        If Ptr = 0 Then
'            ColItemIndex = -1 'key not found
'        End If
'    Else
'        Err.Raise ObjectNotSet 'No object
'    End If
'
'End Function

Public Function GetTempDir() As String
Dim sBuffer As String * 260
Dim lLength As Long

    lLength = GetTempPath(Len(sBuffer), sBuffer)
    GetTempDir = Left$(sBuffer, lLength)
    
    If Right$(GetTempDir, 1) <> "\" Then
        GetTempDir = GetTempDir & "\"
    End If
    
End Function

Public Sub LoadCommands()
Dim bDatabase() As Byte
Dim iFileNum As Integer
Dim i As Integer
Dim j As Integer

    bDatabase = LoadResData(101, "cmddb")
    sTempPath = GetTempDir
    
    iFileNum = FreeFile
    Open sTempPath & "command.dat" For Binary As #iFileNum
        
        Put #iFileNum, 1, bDatabase
        
        ReDim RubiCommands(&HE2) As tRubiCommand
        ReDim RubiParams(&HE2, 5) As tRubiParam
        
        Erase bDatabase
        Seek #iFileNum, 1
        
        Get #iFileNum, , RubiCommands
        
        For i = LBound(RubiCommands) To UBound(RubiCommands)
            For j = 0 To 5
                Get #iFileNum, , RubiParams(i, j)
            Next j
        Next i
        
    Close #iFileNum
    
    DeleteFile sTempPath & "command.dat"
    DeleteFile App.Path & "\command.dat"

End Sub

Public Sub LoadMovements()
Dim i As Integer
    
    ReDim sMoveLabels(0 To 511) As String
    
    For i = LBound(sMoveLabels) To UBound(sMoveLabels)
        sMoveLabels(i) = "mov" & Hex$(i And &HFF)
    Next i
    
    sMoveLabels(&H0) = "Face Down"
    sMoveLabels(&H1) = "Face Up"
    sMoveLabels(&H2) = "Face Left"
    sMoveLabels(&H3) = "Face Right"
    sMoveLabels(&H4) = "Face Down (Faster)"
    sMoveLabels(&H5) = "Face Up (Faster)"
    sMoveLabels(&H6) = "Face Left (Faster)"
    sMoveLabels(&H7) = "Face Right (Faster)"
    sMoveLabels(&H8) = "Step Down (Very Slow)"
    sMoveLabels(&H9) = "Step Up (Very Slow)"
    sMoveLabels(&HA) = "Step Left (Very Slow)"
    sMoveLabels(&HB) = "Step Right (Very Slow)"
    sMoveLabels(&HC) = "Step Down (Slow)"
    sMoveLabels(&HD) = "Step Up (Slow)"
    sMoveLabels(&HE) = "Step Left (Slow)"
    sMoveLabels(&HF) = "Step Right (Slow)"
    sMoveLabels(&H10) = "Step Down (Normal)"
    sMoveLabels(&H11) = "Step Up (Normal)"
    sMoveLabels(&H12) = "Step Left (Normal)"
    sMoveLabels(&H13) = "Step Right (Normal)"
    sMoveLabels(&H14) = "Jump2 Down"
    sMoveLabels(&H15) = "Jump2 Up"
    sMoveLabels(&H16) = "Jump2 Left"
    sMoveLabels(&H17) = "Jump2 Right"
    sMoveLabels(&H18) = "Delay1"
    sMoveLabels(&H19) = "Delay2"
    sMoveLabels(&H1A) = "Delay3"
    sMoveLabels(&H1B) = "Delay4"
    sMoveLabels(&H1C) = "Delay5"
    sMoveLabels(&H1D) = "Step Down (Fast)"
    sMoveLabels(&H1E) = "Step Up (Fast)"
    sMoveLabels(&H1F) = "Step Left (Fast)"
    sMoveLabels(&H20) = "Step Right (Fast)"
    sMoveLabels(&H21) = "Step on the Spot Down (Normal)"
    sMoveLabels(&H22) = "Step on the Spot Up (Normal)"
    sMoveLabels(&H23) = "Step on the Spot Left (Normal)"
    sMoveLabels(&H24) = "Step on the Spot Right (Normal)"
    sMoveLabels(&H25) = "Step on the Spot Down (Faster)"
    sMoveLabels(&H26) = "Step on the Spot Up (Faster)"
    sMoveLabels(&H27) = "Step on the Spot Left (Faster)"
    sMoveLabels(&H28) = "Step on the Spot Right (Faster)"
    sMoveLabels(&H29) = "Step on the Spot Down (Fastest)"
    sMoveLabels(&H2A) = "Step on the Spot Up (Fastest)"
    sMoveLabels(&H2B) = "Step on the Spot Left (Fastest)"
    sMoveLabels(&H2C) = "Step on the Spot Right (Fastest)"
    sMoveLabels(&H2D) = "Face Down (Delayed)"
    sMoveLabels(&H2E) = "Face Up (Delayed)"
    sMoveLabels(&H2F) = "Face Left (Delayed)"
    sMoveLabels(&H30) = "Face Right (Delayed)"
    sMoveLabels(&H31) = "Slide Down (Slow)"
    sMoveLabels(&H32) = "Slide Up (Slow)"
    sMoveLabels(&H33) = "Slide Left (Slow)"
    sMoveLabels(&H34) = "Slide Right (Slow)"
    sMoveLabels(&H35) = "Slide Down (Normal)"
    sMoveLabels(&H36) = "Slide Up (Normal)"
    sMoveLabels(&H37) = "Slide Left (Normal)"
    sMoveLabels(&H38) = "Slide Right (Normal)"
    sMoveLabels(&H39) = "Slide Down (Fast)"
    sMoveLabels(&H3A) = "Slide Up (Fast)"
    sMoveLabels(&H3B) = "Slide Left (Fast)"
    sMoveLabels(&H3C) = "Slide Right (Fast)"
    sMoveLabels(&H3D) = "Slide Running on Right Foot (Down)"
    sMoveLabels(&H3E) = "Slide Running on Right Foot (Up)"
    sMoveLabels(&H3F) = "Slide Running on Right Foot (Left)"
    sMoveLabels(&H40) = "Slide Running on Right Foot (Right)"
    sMoveLabels(&H41) = "Slide Running on Left Foot (Down)"
    sMoveLabels(&H42) = "Slide Running on Left Foot (Up)"
    sMoveLabels(&H43) = "Slide Running on Left Foot (Left)"
    sMoveLabels(&H44) = "Slide Running on Left Foot (Right)"
    sMoveLabels(&H46) = "Jump Facing Left (Down)"
    sMoveLabels(&H47) = "Jump Facing Down (Up)"
    sMoveLabels(&H48) = "Jump Facing Up (Left)"
    sMoveLabels(&H49) = "Jump Facing Left (Right)"
    sMoveLabels(&H4A) = "Face Player"
    sMoveLabels(&H4B) = "Face Against Player"
    sMoveLabels(&H4E) = "Jump Down"
    sMoveLabels(&H4F) = "Jump Up"
    sMoveLabels(&H50) = "Jump Left"
    sMoveLabels(&H51) = "Jump Right"
    sMoveLabels(&H52) = "Jump in Place (Facing Down)"
    sMoveLabels(&H53) = "Jump in Place (Facing Up)"
    sMoveLabels(&H54) = "Jump in Place (Facing Left)"
    sMoveLabels(&H55) = "Jump in Place (Facing Right)"
    sMoveLabels(&H56) = "Jump in Place (Facing Down/Up)"
    sMoveLabels(&H57) = "Jump in Place (Facing Up/Down)"
    sMoveLabels(&H58) = "Jump in Place (Facing Left/Right)"
    sMoveLabels(&H59) = "Jump in Place (Facing Right/Left)"
    sMoveLabels(&H60) = "Hide"
    sMoveLabels(&H61) = "Show"
    sMoveLabels(&H62) = "Exclamation Mark (!)"
    sMoveLabels(&H63) = "Question Mark (?)"
    sMoveLabels(&H64) = "Cross (X)"
    sMoveLabels(&H65) = "Double Exclamation Mark (!!)"
    sMoveLabels(&H66) = "Happy (^_^)"
    
    sMoveLabels(&HFE) = "End of Movements"
    
    sMoveLabels(&H100) = sMoveLabels(&H0)
    sMoveLabels(&H101) = sMoveLabels(&H1)
    sMoveLabels(&H102) = sMoveLabels(&H2)
    sMoveLabels(&H103) = sMoveLabels(&H3)
    sMoveLabels(&H104) = "Step Down (Slow)"
    sMoveLabels(&H105) = "Step Up (Slow)"
    sMoveLabels(&H106) = "Step Left (Slow)"
    sMoveLabels(&H107) = "Step Right (Slow)"
    sMoveLabels(&H108) = "Step Down (Normal)"
    sMoveLabels(&H109) = "Step Up (Normal)"
    sMoveLabels(&H10A) = "Step Left (Normal)"
    sMoveLabels(&H10B) = "Step Right (Normal)"
    sMoveLabels(&H10C) = "Jump2 Down"
    sMoveLabels(&H10D) = "Jump2 Up"
    sMoveLabels(&H10E) = "Jump2 Left"
    sMoveLabels(&H10F) = "Jump2 Right"
    sMoveLabels(&H110) = sMoveLabels(&H18)
    sMoveLabels(&H111) = sMoveLabels(&H19)
    sMoveLabels(&H112) = sMoveLabels(&H1A)
    sMoveLabels(&H113) = sMoveLabels(&H1B)
    sMoveLabels(&H114) = sMoveLabels(&H1C)
    sMoveLabels(&H115) = "Slide Down"
    sMoveLabels(&H116) = "Slide Up"
    sMoveLabels(&H117) = "Slide Left"
    sMoveLabels(&H118) = "Slide Right"
    sMoveLabels(&H119) = "Step on the Spot Down (Slow)"
    sMoveLabels(&H11A) = "Step on the Spot Up (Slow)"
    sMoveLabels(&H11B) = "Step on the Spot Left (Slow)"
    sMoveLabels(&H11C) = "Step on the Spot Right (Slow)"
    sMoveLabels(&H11D) = "Step on the Spot Down (Normal)"
    sMoveLabels(&H11E) = "Step on the Spot Up (Normal)"
    sMoveLabels(&H11F) = "Step on the Spot Left (Normal)"
    sMoveLabels(&H120) = "Step on the Spot Right (Normal)"
    sMoveLabels(&H121) = "Step on the Spot Down (Faster)"
    sMoveLabels(&H122) = "Step on the Spot Up (Faster)"
    sMoveLabels(&H123) = "Step on the Spot Left (Faster)"
    sMoveLabels(&H124) = "Step on the Spot Right (Faster)"
    sMoveLabels(&H125) = "Step on the Spot Down (Fastest)"
    sMoveLabels(&H126) = "Step on the Spot Up (Fastest)"
    sMoveLabels(&H127) = "Step on the Spot Left (Fastest)"
    sMoveLabels(&H128) = "Step on the Spot Right (Fastest)"
    sMoveLabels(&H129) = "Slide Down"
    sMoveLabels(&H12A) = "Slide Up"
    sMoveLabels(&H12B) = "Slide Left"
    sMoveLabels(&H12C) = "Slide Right"
    sMoveLabels(&H12D) = "Slide Down"
    sMoveLabels(&H12E) = "Slide Up"
    sMoveLabels(&H12F) = "Slide Left"
    sMoveLabels(&H130) = "Slide Right"
    sMoveLabels(&H131) = "Slide Down"
    sMoveLabels(&H132) = "Slide Up"
    sMoveLabels(&H133) = "Slide Left"
    sMoveLabels(&H134) = "Slide Right"
    sMoveLabels(&H135) = "Slide Running Down"
    sMoveLabels(&H136) = "Slide Running Up"
    sMoveLabels(&H137) = "Slide Running Left"
    sMoveLabels(&H138) = "Slide Running Right"
    sMoveLabels(&H13A) = "Jump Facing Left (Down)"
    sMoveLabels(&H13B) = "Jump Facing Down (Up)"
    sMoveLabels(&H13C) = "Jump Facing Up (Left)"
    sMoveLabels(&H13D) = "Jump Facing Left (Right)"
    sMoveLabels(&H13E) = sMoveLabels(&H4A)
    sMoveLabels(&H13F) = sMoveLabels(&H4B)
    sMoveLabels(&H142) = sMoveLabels(&H4E)
    sMoveLabels(&H143) = sMoveLabels(&H4F)
    sMoveLabels(&H144) = sMoveLabels(&H50)
    sMoveLabels(&H145) = sMoveLabels(&H51)
    sMoveLabels(&H146) = sMoveLabels(&H52)
    sMoveLabels(&H147) = sMoveLabels(&H53)
    sMoveLabels(&H148) = sMoveLabels(&H54)
    sMoveLabels(&H149) = sMoveLabels(&H55)
    sMoveLabels(&H14A) = sMoveLabels(&H56)
    sMoveLabels(&H14B) = sMoveLabels(&H57)
    sMoveLabels(&H14C) = sMoveLabels(&H58)
    sMoveLabels(&H14D) = sMoveLabels(&H59)
    sMoveLabels(&H14E) = "Face Left"
    sMoveLabels(&H154) = sMoveLabels(&H60)
    sMoveLabels(&H155) = sMoveLabels(&H61)
    sMoveLabels(&H156) = sMoveLabels(&H62)
    sMoveLabels(&H157) = sMoveLabels(&H63)
    sMoveLabels(&H158) = "Love (<3)"
    sMoveLabels(&H162) = "Walk Down"
    sMoveLabels(&H163) = "Walk Down"
    sMoveLabels(&H164) = sMoveLabels(&H2D)
    sMoveLabels(&H165) = sMoveLabels(&H2E)
    sMoveLabels(&H166) = sMoveLabels(&H2F)
    sMoveLabels(&H167) = sMoveLabels(&H30)
    sMoveLabels(&H170) = sMoveLabels(&H52)
    sMoveLabels(&H171) = sMoveLabels(&H53)
    sMoveLabels(&H172) = sMoveLabels(&H54)
    sMoveLabels(&H173) = sMoveLabels(&H55)
    sMoveLabels(&H174) = "Jump Down Running"
    sMoveLabels(&H175) = "Jump Up Running"
    sMoveLabels(&H176) = "Jump Left Running"
    sMoveLabels(&H177) = "Jump Right Running"
    sMoveLabels(&H178) = "Jump2 Down Running"
    sMoveLabels(&H179) = "Jump2 Up Running"
    sMoveLabels(&H17A) = "Jump2 Left Running"
    sMoveLabels(&H17B) = "Jump2 Right Running"
    sMoveLabels(&H17C) = "Walk on the Spot (Down)"
    sMoveLabels(&H17D) = "Walk on the Spot (Up)"
    sMoveLabels(&H17E) = "Walk on the Spot (Lef)"
    sMoveLabels(&H17F) = "Walk on the Spot (Right)"
    sMoveLabels(&H180) = "Slide Down Running"
    sMoveLabels(&H181) = "Slide Up Running"
    sMoveLabels(&H182) = "Slide Left Running"
    sMoveLabels(&H183) = "Slide Right Running"
    sMoveLabels(&H184) = "Slide Down"
    sMoveLabels(&H185) = "Slide Up"
    sMoveLabels(&H186) = "Slide Left"
    sMoveLabels(&H187) = "Slide Right"
    sMoveLabels(&H188) = "Slide Down on Left Foot"
    sMoveLabels(&H189) = "Slide Up on Left Foot"
    sMoveLabels(&H18A) = "Slide Left on Left Foot "
    sMoveLabels(&H18B) = "Slide Right on Left Foot"
    sMoveLabels(&H18C) = "Slide Left diagonally (Facing Up)"
    sMoveLabels(&H18D) = "Slide Right diagonally (Facing Up)"
    sMoveLabels(&H18E) = "Slide Left diagonally (Facing Down)"
    sMoveLabels(&H18F) = "Slide Right diagonally (Facing Down)"
    sMoveLabels(&H190) = "Slide2 Left diagonally (Facing Up)"
    sMoveLabels(&H191) = "Slide2 Right diagonally (Facing Up)"
    sMoveLabels(&H192) = "Slide2 Left diagonally (Facing Down)"
    sMoveLabels(&H193) = "Slide2 Right diagonally (Facing Down)"
    sMoveLabels(&H196) = "Walk Left"
    sMoveLabels(&H197) = "Walk Right"
    sMoveLabels(&H198) = "Levitate"
    sMoveLabels(&H199) = "Stop Levitating"
    sMoveLabels(&H19C) = "Fly Up Vertically"
    sMoveLabels(&H19D) = "Land"
    
    sMoveLabels(&H1FE) = sMoveLabels(&HFE)
    
    MovesReady = True
     
End Sub

Public Function GetMovementLabel(bMovement As Byte) As String
Dim iLabelIndex As Integer
    
    If MovesReady = False Then
        LoadMovements
    End If
    
    iLabelIndex = bMovement
    
    If IsFireRed = False Then
        iLabelIndex = iLabelIndex + &H100
    End If
    
    GetMovementLabel = sMoveLabels(iLabelIndex)

End Function

Private Function CByt(ByRef sValue As String) As Byte
Dim lErr As Long

    On Error GoTo ErrHandler
    CByt = CByte(sValue)
    
    Exit Function
    
ErrHandler:
    
    lErr = Err.Number
    
    If lErr = TypeMismatch Then
        If InStrB(sValue, "&H") = 0 Then
            MissingDefine = True
        End If
    End If
    
    On Error GoTo 0
    Err.Clear
    Err.Raise lErr
    
End Function

Private Function CWord(ByRef sValue As String) As Integer
Dim lErr As Long

    On Error GoTo ErrHandler
    CWord = CInt(sValue)
    
    Exit Function
    
ErrHandler:
    
    lErr = Err.Number
    
    If lErr = TypeMismatch Then
        If InStrB(sValue, "&H") = 0 Then
            MissingDefine = True
        End If
    End If
    
    On Error GoTo 0
    Err.Clear
    Err.Raise lErr
    
    
End Function

Private Function CLong(ByRef sValue As String) As Long
Dim lErr As Long

    On Error GoTo ErrHandler
    CLong = CLng(sValue)
    
    Exit Function
    
ErrHandler:
    
    lErr = Err.Number
    
    If lErr = TypeMismatch Then
        If InStrB(sValue, "&H") = 0 Then
            MissingDefine = True
        End If
    End If
    
    On Error GoTo 0
    Err.Clear
    Err.Raise lErr
     
End Function

Public Function CPtr(ByRef sOffset As String) As Long
Dim lOffset As Long
Dim lErr As Long
    
    On Error GoTo ErrHandler
    
    If AscW(sOffset) <> ascAt Then
        lOffset = CLng(sOffset)
    Else
        lOffset = BinarySearchDefine(DynamicOffsets2, sOffset)
    End If
    
    If lOffset > 0 Then

        If lOffset <= &H9FFFFFF Then
        
            If AutoBank = True Then
            
                If lOffset <= &H1FFFFFF Then
                    CPtr = lOffset + &H8000000
                Else
                    CPtr = lOffset
                End If
            
            Else
                CPtr = lOffset
            End If
        
        Else
            Err.Raise Overflow
        End If
        
    ElseIf lOffset = 0 Then
        CPtr = lOffset
    Else
        Err.Raise TypeMismatch
    End If
    
    Exit Function
    
ErrHandler:
    
    lErr = Err.Number
    
    If lErr = TypeMismatch Then
        If InStrB(sOffset, "&H") = 0 Then
            If AscW(sOffset) = ascAt Then
                MissingDynamic = True
            Else
                MissingDefine = True
            End If
        End If
    End If
    
    On Error GoTo 0
    Err.Clear
    Err.Raise lErr
    
End Function

Public Sub ClearData()
    EraseCol Defines
    EraseCol Headers
    EraseCol Aliases
    Erase DynamicOffsets2
End Sub

Private Sub AddDefine(sSymbol As String, sValue As String)
    
    On Error GoTo AlreadyIn
    Defines.Add sValue, sSymbol
    
AlreadyIn:
End Sub

Private Function RemoveDefine(sSymbol As String)
    
    On Error GoTo NotIn
    Defines.Remove sSymbol
    
NotIn:
End Function

Private Function AddHeader(sFile As String) As Boolean
    
    On Error GoTo AlreadyIn
    
    Headers.Add sFile, sFile
    AddHeader = True
    
AlreadyIn:
End Function

Private Sub AddAlias(sSymbol As String, sNewSymbol As String)
    
    On Error GoTo AlreadyIn
    Aliases.Add sNewSymbol, sSymbol
    
AlreadyIn:
End Sub

Private Function RemoveAlias(sSymbol As String)
    
    On Error GoTo NotIn
    Defines.Remove sSymbol
    
NotIn:
End Function

Public Sub HeaderProcess(ByVal sFileName As String)
'Const vbDoubleSpace As String = "  "
Dim sRawInput As String
Dim sTempInput As String
Dim sTemp As String
Dim BlockComment As Boolean
Dim sKeywords() As String
Dim lKeyCount As Long
Dim iFileNum As Integer
Dim lTemp As Long
Dim sSingleComments(2) As String
Dim i As Long
    
    If AddHeader(GetFileName(sFileName)) = False Then
        Exit Sub
    End If
    
    If InStrB(sFileName, "\") = 0 Then
        sFileName = App.Path & "\" & sFileName
    End If
    
    If FileExists(sFileName) = False Then
        Err.Raise FileNotFound
        Exit Sub
    End If
    
    ReDim sKeywords(MaxKeywords) As String
    
    sSingleComments(0) = "'"
    sSingleComments(1) = ";"
    sSingleComments(2) = "//"

    iFileNum = FreeFile
    Open sFileName For Input As #iFileNum
        
        Do
    
NextLine:
    
            If EOF(iFileNum) Then Exit Do
            
            Line Input #iFileNum, sRawInput
                    
            If LenB(sRawInput) <> 0 Then
            
                For i = LBound(sSingleComments) To UBound(sSingleComments)
                
                    lTemp = InStrB(sRawInput, sSingleComments(i))
                    
                    If lTemp <> 0 Then
                    
                        If InStrB(sRawInput, "=") <> 1& Then
                            
                            sRawInput = RTrim$(MidB$(sRawInput, 1, lTemp - 1&))
                            
                            If LenB(sRawInput) = 0 Then
                                GoTo NextLine
                            End If
                            
                        Else
                            
                            If lTemp = 1 Then
                                GoTo NextLine
                            End If
                            
                        End If
                        
                    End If
                
                Next i
                
                If InStrB(sRawInput, "=") <> 1 Then
                    sRawInput = RTrim$(sRawInput)
                End If
            
            Else
                GoTo NextLine
            End If
            
            If LenB(sRawInput) <> 0 Then
            
                lTemp = InStrB(sRawInput, "/*")
                
                If lTemp <> 0 Then
                    
                    i = 0
                    
                    sTempInput = MidB$(sRawInput, 1, lTemp - 1)
                    
                    If InStrB(sTempInput, "=") <> 1 Then
                        sTempInput = RTrim$(sTempInput)
                    End If
                        
                    Do While InStrB(sRawInput, "*/") = 0
                        
                        i = i + 1
                        If EOF(iFileNum) Then Exit Do
                        
                        Line Input #iFileNum, sRawInput
                        
                        If InStrB(sRawInput, "*/") <> 0 Then
                            Exit Do
                        End If
                        
                    Loop
                    
                    If i > 0 Then
                
                        lTemp = InStrB(sRawInput, "*/")
                        
                        If lTemp <> 0 Then
                            
                            If LenB(sTempInput) <> 0 Then
                                BlockComment = True
                            End If
                            
                            sRawInput = MidB$(sRawInput, lTemp + 4)
                            
                            If InStrB(sRawInput, "=") <> 1 Then
                                sRawInput = RTrim$(sTempInput)
                            End If
                            
                        Else
                            sRawInput = sTempInput
                        End If
                    
                    Else
                        sRawInput = sTempInput
                    End If
                    
                End If
                
                If LenB(sRawInput) <> 0 Then
                
                    For i = LBound(sSingleComments) To UBound(sSingleComments)
                    
                        lTemp = InStrB(sRawInput, sSingleComments(i))
                        
                        If lTemp <> 0 Then
                        
                            If InStrB(1, sRawInput, "=") <> 1& Then
                                
                                sRawInput = RTrim$(MidB$(sRawInput, 1, lTemp - 1&))
                                
                                If LenB(sRawInput) = 0 Then
                                    GoTo NextLine
                                End If
                                
                            Else
                                
                                If lTemp = 1 Then
                                    GoTo NextLine
                                End If
                                
                            End If
                            
                        End If
                    
                    Next i
                
                Else
                    GoTo NextLine
                End If
                
            Else
                GoTo NextLine
            End If
    
            If BlockComment Then
                sTemp = sRawInput
                sRawInput = sTempInput
            End If

Continue:

            sRawInput = LCase$(sRawInput)
            SplitB sRawInput, sKeywords(), vbSpace, lKeyCount
            
            If lKeyCount > 2& Then
    
                Select Case sKeywords(0)
                    Case "#define", "#const"
                        AddDefine sKeywords(1), sKeywords(2)
                    Case "#alias"
                        AddAlias sKeywords(1), sKeywords(2)
                End Select
                
            End If
                
            If BlockComment Then
                sRawInput = sTemp
                BlockComment = False
                GoTo Continue
            End If
                
        Loop
    
    Close #iFileNum
  
End Sub

Public Function Process(sFileName As String, Optional sScriptFile As String = vbNullString, Optional Batch As Boolean = False) As Boolean
Dim iDestFile As Integer
Dim iFileNum As Integer
Dim sRawInput As String
Dim sTempInput As String
Dim AlreadyAdded As Boolean
Dim sSingleComments(2) As String
Dim sKeywords() As String
Dim lKeyCount As Long
Dim lParamCount As Long
Dim lLineNo As Long
Dim lLineCount As Long
Dim sTemp As String
Dim i As Long
Dim j As Long
Dim sErrorDescr As String
Dim bTemp As Byte
Dim bTempArray() As Byte
Dim bTempArray2() As Byte
Dim sArray() As String
Dim lTemp As Long
Dim FirstOrg As Boolean
Dim lOrgCount As Long
Dim lCurrentOrg As Long
Dim lDynamicStart As Long
Dim DynamicOffsets() As Define
Dim Defines2() As Define
Dim TextData() As ByteString
Dim lTextCount As Long
Dim lTotalBytes() As Long
Dim colTest As Collection
Const lChunkSize As Long = &H10000 '64K
Dim lMaxFileSize As Long
Dim bBuffer() As Byte
Dim bSearch() As Byte
Dim lFoundOffset As Long
Dim cLog As cStringBuilder
Dim cTemp As cStringBuilder
Dim PreProcessed As Boolean

    If FileExists(sFileName) = False Then
        MsgBox LoadResString(13001), vbExclamation
        Exit Function
    End If
    
    StartTiming
    
    On Error GoTo StdErr
    
    Process = True
    AutoBank = True
    FirstOrg = True
    lDynamicCount = 0
    lDynamicStart = -1
    ReDim TextData(0) As ByteString
    ReDim BrailleData(0) As ByteString
    ReDim DynamicOffsets(0) As Define
    ReDim lTotalBytes(0) As Long
    Set colTest = New Collection
    Set cTemp = New cStringBuilder
    Set cLog = New cStringBuilder
    
    ClearData
    
    If FileExists(App.Path & "\std.rbh") Then
        HeaderProcess App.Path & "\std.rbh"
    End If
    
    If IsLookupReady = False Then
        
        ReDim RubiLookup(UBound(RubiCommands)) As Define
        
        For i = LBound(RubiLookup) To UBound(RubiLookup)
            RubiLookup(i).Symbol = RubiCommands(i).Keyword
            RubiLookup(i).Value = i
        Next i
        
        RubiLookup(&H5C).Symbol = "trnbattle"
        
        TriQuickSortDefine RubiLookup
        
    End If
    
    MakeWritable sFileName
    
    iDestFile = FreeFile
    Open sFileName For Binary As #iDestFile
        
        If LOF(iDestFile) <> 0 Then
            lMaxFileSize = LOF(iDestFile)
        Else
            If IsDebugging Then
                lMaxFileSize = &H1000000
            End If
        End If
    
        Get #iDestFile, &HAD, sGameCode
        bFreeSpace = &HFF
        
        Select Case sGameCode
            Case "AXV", "AXP"
                MaxCommand = &HC5
            Case "BPR", "BPG"
                MaxCommand = &HD4
            Case Else
                MaxCommand = &HE2
        End Select
        
        If IsDebugging Then
            bFreeSpace = 0
        End If
        
If sGameCode = "BPE" Then
bFreeSpace = &HFF
End If
        
        sSingleComments(0) = "'"
        sSingleComments(1) = ";"
        sSingleComments(2) = "//"
        
        iFileNum = FreeFile
        
        If LenB(sScriptFile) = 0 Then
            Open sTempPath & sTempFile For Input As #iFileNum
        Else
            Open sScriptFile For Input As #iFileNum
        End If
        
Begin:
        If Not PreProcessed Then
            cLog.Append App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine
            cLog.Append String$(37, "-") & vbNewLine
            cLog.Append Date$ & vbSpace & TimeValue(time$) & vbNewLine
            cLog.Append String$(37, "-") & vbNewLine
            cLog.Append LoadResString(13002) & sFileName & "..." & vbNewLine
            cLog.Append LoadResString(13003) & vbNewLine
        End If
              
        Do
            
            If Not PreProcessed Then

NextLine:
                If EOF(iFileNum) Then Exit Do
                
                lLineNo = lLineNo + 1&
                Line Input #iFileNum, sRawInput
                
                If LenB(sRawInput) <> 0 Then
                
                    For i = LBound(sSingleComments) To UBound(sSingleComments)
                    
                        lTemp = InStrB(sRawInput, sSingleComments(i))
                        
                        If lTemp <> 0 Then
                        
                            If InStrB(1, sRawInput, "=") <> 1& Then
                                
                                sRawInput = RTrim$(MidB$(sRawInput, 1, lTemp - 1&))
                                
                                If LenB(sRawInput) = 0 Then
                                    cTemp.Append vbNewLine
                                    GoTo NextLine
                                End If
                                
                            Else
                                
                                If lTemp = 1 Then
                                    cTemp.Append vbNewLine
                                    GoTo NextLine
                                End If
                                
                            End If
                            
                        End If
                    
                    Next i
                    
                    If InStrB(1, sRawInput, "=") <> 1 Then
                        sRawInput = RTrim$(sRawInput)
                    End If
                
                Else
                    cTemp.Append vbNewLine
                    GoTo NextLine
                End If
                
                If LenB(sRawInput) <> 0 Then
                
                    lTemp = InStrB(sRawInput, "/*")
                    
                    If lTemp <> 0 Then
                        
                        j = 0
                        sTempInput = MidB$(sRawInput, 1, lTemp - 1&)
                        
                        If InStrB(sTempInput, "=") <> 1 Then
                            sTempInput = RTrim$(sTempInput)
                        End If
                            
                        Do While InStrB(sRawInput, "*/") = 0
                            
                            j = j + 1&
                            If EOF(iFileNum) Then Exit Do
                            
                            Line Input #iFileNum, sRawInput
                            lLineNo = lLineNo + 1&
                            
                            If InStrB(sRawInput, "*/") <> 0 Then
                                Exit Do
                            End If
                            
                        Loop
                        
                        If j > 0 Then
                    
                            lTemp = InStrB(sRawInput, "*/")
                            
                            If lTemp <> 0 Then
                                
                                If LenB(sTempInput) <> 0 Then
                                    AlreadyAdded = True
                                    cTemp.Append sTempInput & vbNewLine
                                Else
                                    cTemp.Append vbNewLine
                                End If
                                
                                For j = 1 To j - 1&
                                    cTemp.Append vbNewLine
                                Next j
                                
                                sRawInput = MidB$(sRawInput, lTemp + 4&)
                                
                                If InStrB(1, sRawInput, "=") <> 1 Then
                                    sRawInput = RTrim$(sRawInput)
                                End If
                                
                            Else
                                sRawInput = sTempInput
                            End If
                        
                        Else
                            sRawInput = sTempInput
                        End If
                        
                    End If
                    
                    If LenB(sRawInput) <> 0 Then
                    
                        For i = LBound(sSingleComments) To UBound(sSingleComments)
                        
                            lTemp = InStrB(sRawInput, sSingleComments(i))
                            
                            If lTemp <> 0 Then
                            
                                If InStrB(1, sRawInput, "=") <> 1& Then
                                    
                                    sRawInput = RTrim$(MidB$(sRawInput, 1, lTemp - 1&))
                                    
                                    If LenB(sRawInput) = 0 Then
                                        cTemp.Append vbNewLine
                                        GoTo NextLine
                                    End If
                                    
                                Else
                                    
                                    If lTemp = 1 Then
                                        cTemp.Append vbNewLine
                                        GoTo NextLine
                                    End If
                                    
                                End If
                                
                            End If
                        
                        Next i
                    
                    Else
                        cTemp.Append vbNewLine
                        GoTo NextLine
                    End If
                    
                Else
                    cTemp.Append vbNewLine
                    GoTo NextLine
                End If
                
            Else
            
NextOne:
                sRawInput = sArray(lLineNo)
                lLineNo = lLineNo + 1&
                
                If lLineNo = lLineCount Then
                    Exit Do
                End If
                
                If LenB(sRawInput) = 0 Then
                    GoTo NextOne
                End If
                
            End If
                                                    
            If InStrB(sRawInput, "=") <> 1 Then
                
                If Not PreProcessed Then
                    
                    sRawInput = LCase$(sRawInput)
                    cTemp.Append sRawInput & vbNewLine
                    
                    If AlreadyAdded Then
                        sTemp = sRawInput
                        sRawInput = sTempInput
                    End If
                    
                    If Aliases.Count > 0 Then
                        For i = 1& To Aliases.Count
                            DoReplace sRawInput, Aliases.Item(i), ColItemKey(i, Aliases)
                        Next i
                    End If
                
                Else
                    DoReplace sRawInput, "0x", "&H"
                End If
                
                SplitB sRawInput, sKeywords(), vbSpace, lKeyCount
                
            Else
                
                sKeywords(0) = "="
                cTemp.Append sRawInput & vbNewLine
                
                If Not PreProcessed Then
                    If AlreadyAdded Then
                        sTemp = sRawInput
                        sRawInput = sTempInput
                    End If
                End If
                
            End If
            
Continue:
                
            If InStrB(sKeywords(0), "#") <> 1 Then
                 
                i = BinarySearchDefine(RubiLookup, sKeywords(0))
                 
                If i > -1& Then
                     
                    If lOrgCount = 0 Then GoTo NoOrg
                    If i > MaxCommand Then GoTo InvalidCommand
                     
                    If Not PreProcessed Then
                     
                        lParamCount = RubiCommands(i).ParamCount
                         
                        If lKeyCount < lParamCount + 1& Then
                            GoTo TooLessParams
                        ElseIf lKeyCount > lParamCount + 1& Then
                            GoTo TooMuchParams
                        End If
                         
                        lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + RubiCommands(i).NeededBytes
                     
                    Else
                         
                        If Loc(iDestFile) <= lMaxFileSize - RubiCommands(i).NeededBytes Then
                            Put #iDestFile, , CByte(i)
                        Else
                            Err.Raise BadRecordNumber
                        End If
                                                                         
                        cLog.Append lLineNo & " - (" & RightB$("0" & Hex$(i), 4&) & ") - " & UCase$(RubiCommands(i).Keyword) & " [+" & RubiCommands(i).NeededBytes & "]" & vbNewLine
                        
                        If RubiCommands(i).ParamCount > 0 Then
                            For j = 0& To RubiCommands(i).ParamCount - 1
                                Select Case RubiParams(i, j).SIZE
                                    Case 1
                                        Put #iDestFile, , CWord(sKeywords(1& + j))
                                        cLog.Append " > iWord =" & sHexPrefix & Hex$(CInt(sKeywords(1& + j))) & vbNewLine
                                    Case 3
                                        Put #iDestFile, , CPtr(sKeywords(1& + j))
                                        cLog.Append " > pPointer =" & sHexPrefix & Hex$(CPtr(sKeywords(1& + j))) & vbNewLine
                                    Case 0
                                        Put #iDestFile, , CByt(sKeywords(1& + j))
                                        cLog.Append " > bByte =" & sHexPrefix & Hex$(CByte(sKeywords(1& + j))) & vbNewLine
                                    Case 2
                                        Put #iDestFile, , CLong(sKeywords(1& + j))
                                        cLog.Append " > lDword =" & sHexPrefix & Hex$(CLng(sKeywords(1& + j))) & vbNewLine
                                End Select
                            Next j
                        End If
                     
                    End If
                             
                Else
                     
                    Select Case sKeywords(0)
                 
                        Case "="
                            
                            If Not PreProcessed Then
                            
                                If lOrgCount = 0 Then
                                    GoTo NoOrg
                                End If
                                
                                TextData(lTextCount).sText = MidB$(sRawInput, 5&)
                                
                                lTemp = Asc2SappLen(TextData(lTextCount).sText) + 1&
                                ReDim TextData(lTextCount).bArray(lTemp - 1) As Byte
                                
                                Asc2Sapp TextData(lTextCount).sText & "\x", TextData(lTextCount).bArray
                                
                                If lTextCount = UBound(TextData) Then
                                    ReDim Preserve TextData(lTextCount + 19&) As ByteString
                                End If
                                
                                lTextCount = lTextCount + 1&
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + lTemp + 1&
                                
                            Else
                                
                                If Loc(iDestFile) <= lMaxFileSize - (UBound(TextData(lTextCount).bArray) + 1&) Then
                                    Put #iDestFile, , TextData(lTextCount).bArray
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                                
                                If LenB(DynamicOffsets(lCurrentOrg).Symbol) <> 0 Then
                                    Put #iDestFile, , CByte(&H0)
                                End If
                                
                                cLog.Append lLineNo & " - RAW TEXT [+" & (UBound(TextData(lTextCount).bArray) + 1&) & "]" & vbNewLine
                                cLog.Append " > sText = """ & TextData(lTextCount).sText & """" & vbNewLine
                                
                                lTextCount = lTextCount + 1&
    
                            End If
                         
                        Case "msgbox", "message"
                            
                            If Not PreProcessed Then
                            
                                If lOrgCount = 0 Then GoTo NoOrg
                                
                                lParamCount = 2&
                                
                                If lKeyCount < lParamCount + 1& Then
                                    GoTo TooLessParams
                                ElseIf lKeyCount > lParamCount + 1& Then
                                    GoTo TooMuchParams
                                End If
                                
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 8&
                                                                    
                            Else
                                                                               
                                If Loc(iDestFile) <= lMaxFileSize - 8& Then
                                    Put #iDestFile, , CByte(&HF)
                                    Put #iDestFile, , CByte(0)
                                    Put #iDestFile, , CPtr(sKeywords(1))
                                    Put #iDestFile, , CByte(&H9)
                                    Put #iDestFile, , CByt(sKeywords(2))
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                                
                                cLog.Append lLineNo & " - (0F) - " & UCase$(sKeywords(0)) & " (native) [+8]" & vbNewLine
                                cLog.Append " > pText =" & sHexPrefix & Hex$(CPtr(sKeywords(1))) & vbNewLine
                                cLog.Append " > bType =" & sHexPrefix & Hex$(CByte(sKeywords(2))) & vbNewLine
                         
                            End If
                         
                        Case "if"
                            
                            If Not PreProcessed Then
                                
                                If lOrgCount = 0 Then GoTo NoOrg
                                
                                lParamCount = 256&
                                
                                If lKeyCount < 2& Then
                                    GoTo TooLessParams
                                End If
                                
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 6&
                                
                            Else
                            
                                If Loc(iDestFile) > lMaxFileSize - 6& Then
                                    Err.Raise BadRecordNumber
                                End If
                                
                                cLog.Append lLineNo & " - (??) - " & UCase$(sKeywords(0)) & " (native) [+6]" & vbNewLine
                                cLog.Append " > bCondition =" & sHexPrefix & Hex$(CByt(sKeywords(1))) & vbNewLine
                                
                                Select Case sKeywords(2)
                                    
                                    Case "goto", "jump"
                                    
                                        Put #iDestFile, , CByte(&H6)
                                        Put #iDestFile, , CByt(sKeywords(1))
                                        Put #iDestFile, , CPtr(sKeywords(3))
                                        
                                        cLog.Append LoadResString(13011) & vbNewLine
                                        cLog.Append " > pTarget =" & sHexPrefix & Hex$(CPtr(sKeywords(3))) & vbNewLine
                                        
                                    Case "call", "gosub"
                                        
                                        Put #iDestFile, , CByte(&H7)
                                        Put #iDestFile, , CByt(sKeywords(1))
                                        Put #iDestFile, , CPtr(sKeywords(3))
                                        
                                        cLog.Append LoadResString(13010) & vbNewLine
                                        cLog.Append " > pTarget =" & sHexPrefix & Hex$(CPtr(sKeywords(3))) & vbNewLine
                                        
                                    Case Else
                                        
                                        Put #iDestFile, , CByte(&H7)
                                        Put #iDestFile, , CByt(sKeywords(1))
                                        Put #iDestFile, , CPtr(sKeywords(2))
                                        
                                        cLog.Append LoadResString(13012) & vbNewLine
                                        cLog.Append " > pTarget =" & sHexPrefix & Hex$(CPtr(sKeywords(2))) & vbNewLine
                                        
                                End Select
                                
                            End If
                             
                        Case "trainerbattle"
                            
                            If Not PreProcessed Then
                            
                                If lOrgCount = 0 Then GoTo NoOrg
                                
                                DoReplace sKeywords(1), "0x", "&H"
                                
                                Select Case CByt(sKeywords(1))
                                    Case &H0, &H5, &H9
                                        lParamCount = 5
                                    Case &H1, &H2, &H4, &H7
                                        lParamCount = 6
                                    Case &H3
                                        lParamCount = 4
                                    Case &H6, &H8
                                        lParamCount = 7
                                    Case Else
                                        lParamCount = 5
                                End Select
                                
                                If lKeyCount < lParamCount + 1& Then
                                    GoTo TooLessParams
                                ElseIf lKeyCount > lParamCount + 1& Then
                                    GoTo TooMuchParams
                                End If
                                
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 6& + ((lKeyCount - 4&) * 4&)
                                
                            Else
                            
                                cLog.Append lLineNo & " - (5C) - " & UCase$(sKeywords(0)) & " [+" & (6 + (lKeyCount - 4&) * 4&) & "]" & vbNewLine
                                cLog.Append " > bKind =" & sHexPrefix & Hex$(CByt(sKeywords(1))) & vbNewLine
                                cLog.Append " > iIndex =" & sHexPrefix & Hex$(CWord(sKeywords(2))) & vbNewLine
                                cLog.Append " > iReserved =" & sHexPrefix & Hex$(CWord(sKeywords(3))) & vbNewLine
                                
                                If CByt(sKeywords(1)) <> 3 Then
                                    cLog.Append " > pChallenge =" & sHexPrefix & Hex$(CPtr(sKeywords(4))) & vbNewLine
                                    cLog.Append " > pDefeat =" & sHexPrefix & Hex$(CPtr(sKeywords(5))) & vbNewLine
                                Else
                                    cLog.Append " > pDefeat =" & sHexPrefix & Hex$(CPtr(sKeywords(4))) & vbNewLine
                                End If
                                
                                If Loc(iDestFile) <= lMaxFileSize - 10& Then
                                    Put #iDestFile, , CByte(&H5C)
                                    Put #iDestFile, , CByt(sKeywords(1))
                                    Put #iDestFile, , CWord(sKeywords(2))
                                    Put #iDestFile, , CWord(sKeywords(3))
                                    Put #iDestFile, , CPtr(sKeywords(4))
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                                
                                If CByte(sKeywords(1)) <> 3 Then
                                    
                                    If Loc(iDestFile) <= lMaxFileSize - 4& Then
                                        Put #iDestFile, , CPtr(sKeywords(5))
                                    Else
                                        Err.Raise BadRecordNumber
                                    End If
                                    
                                    If lKeyCount > 6& Then

                                        cLog.Append " > pSpecial =" & sHexPrefix & Hex$(CPtr(sKeywords(6))) & vbNewLine
                                        
                                        If Loc(iDestFile) <= lMaxFileSize - 4& Then
                                            Put #iDestFile, , CPtr(sKeywords(6))
                                        Else
                                            Err.Raise BadRecordNumber
                                        End If
                                        
                                        If lKeyCount = 8& Then

                                            cLog.Append " > pSpecial2 =" & sHexPrefix & Hex$(CPtr(sKeywords(7))) & vbNewLine
                                            
                                            If Loc(iDestFile) <= lMaxFileSize - 4& Then
                                                Put #iDestFile, , CPtr(sKeywords(7))
                                            Else
                                                Err.Raise BadRecordNumber
                                            End If
                                            
                                        End If
                                        
                                    End If
                                    
                                End If
                                
                            End If
                             
                        Case "giveitem"
                            
                            If Not PreProcessed Then
                            
                                If lOrgCount = 0 Then GoTo NoOrg
                                
                                lParamCount = 3&
                                
                                If lKeyCount < lParamCount + 1& Then
                                    GoTo TooLessParams
                                ElseIf lKeyCount > lParamCount + 1& Then
                                    GoTo TooMuchParams
                                End If
                                
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 12&
                                
                            Else
                                
                                If Loc(iDestFile) <= lMaxFileSize - 12& Then
                                    Put #iDestFile, , CByte(&H1A)
                                    Put #iDestFile, , CInt(&H8000)
                                    Put #iDestFile, , CWord(sKeywords(1))
                                    Put #iDestFile, , CByte(&H1A)
                                    Put #iDestFile, , CInt(&H8001)
                                    Put #iDestFile, , CWord(sKeywords(2))
                                    Put #iDestFile, , CByte(&H9)
                                    Put #iDestFile, , CByt(sKeywords(3))
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                                
                                cLog.Append lLineNo & " - (--) - " & UCase$(sKeywords(0)) & " [+12]" & vbNewLine
                                cLog.Append " > iItem =" & sHexPrefix & Hex$(CInt(sKeywords(1))) & vbNewLine
                                cLog.Append " > iQuantity =" & sHexPrefix & Hex$(CInt(sKeywords(2))) & vbNewLine
                                cLog.Append " > bType =" & sHexPrefix & Hex$(CByte(sKeywords(3))) & vbNewLine
                        
                            End If
                             
                        Case "wildbattle"
                            
                            If Not PreProcessed Then
                            
                                If lOrgCount = 0 Then GoTo NoOrg
                                
                                lParamCount = 3&
                                
                                If lKeyCount < lParamCount + 1& Then
                                    GoTo TooLessParams
                                ElseIf lKeyCount > lParamCount + 1& Then
                                    GoTo TooMuchParams
                                End If
                                
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 7&
                                                                    
                            Else
                                
                                If Loc(iDestFile) <= lMaxFileSize - 7& Then
                                    Put #iDestFile, , CByte(&HB6)
                                    Put #iDestFile, , CWord(sKeywords(1))
                                    Put #iDestFile, , CByt(sKeywords(2))
                                    Put #iDestFile, , CWord(sKeywords(3))
                                    Put #iDestFile, , CByte(&HB7)
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                                
                                cLog.Append lLineNo & " - (--) - " & UCase$(sKeywords(0)) & " [+7]" & vbNewLine
                                cLog.Append " > iSpecies =" & sHexPrefix & Hex$(CInt(sKeywords(1))) & vbNewLine
                                cLog.Append " > bLevel =" & sHexPrefix & Hex$(CByte(sKeywords(2))) & vbNewLine
                                cLog.Append " > iHeldItem =" & sHexPrefix & Hex$(CInt(sKeywords(3))) & vbNewLine

                            End If
                 
                        Case "giveitem2"
                            
                            If Not PreProcessed Then
                            
                                If lOrgCount = 0 Then GoTo NoOrg
                                
                                lParamCount = 3&
                                
                                If lKeyCount < lParamCount + 1& Then
                                    GoTo TooLessParams
                                ElseIf lKeyCount > lParamCount + 1& Then
                                    GoTo TooMuchParams
                                End If
                                
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 17&
                                                            
                            Else
                                
                                If Loc(iDestFile) <= lMaxFileSize - 17& Then
                                    Put #iDestFile, , CByte(&H1A)
                                    Put #iDestFile, , CInt(&H8000)
                                    Put #iDestFile, , CWord(sKeywords(1))
                                    Put #iDestFile, , CByte(&H1A)
                                    Put #iDestFile, , CInt(&H8001)
                                    Put #iDestFile, , CWord(sKeywords(2))
                                    Put #iDestFile, , CByte(&H1A)
                                    Put #iDestFile, , CInt(&H8002)
                                    Put #iDestFile, , CWord(sKeywords(3))
                                    Put #iDestFile, , CInt(&H909)
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                                
                                cLog.Append lLineNo & " - (--) - " & UCase$(sKeywords(0)) & " [+17]" & vbNewLine
                                cLog.Append " > iItem =" & sHexPrefix & Hex$(CInt(sKeywords(1))) & vbNewLine
                                cLog.Append " > iQuantity =" & sHexPrefix & Hex$(CInt(sKeywords(2))) & vbNewLine
                                cLog.Append " > iSoundEffect =" & sHexPrefix & Hex$(CInt(sKeywords(3))) & vbNewLine

                            End If
                                                             
'                        Case "boxset"
'
'                            If Not PreProcessed Then
'
'                                If lOrgCount = 0 Then GoTo NoOrg
'
'                                lParamCount = 1&
'
'                                If lKeyCount < lParamCount + 1& Then
'                                    GoTo TooLessParams
'                                ElseIf lKeyCount > lParamCount + 1& Then
'                                    GoTo TooMuchParams
'                                End If
'
'                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 2&
'
'                            Else
'
'                                If Loc(iDestFile) <= lMaxFileSize - 2 Then
'                                    Put #iDestFile, , CByte(&H9)
'                                    Put #iDestFile, , CByt(sKeywords(1))
'                                Else
'                                    Err.Raise BadRecordNumber
'                                End If
'
'                                cLog.Append lLineNo & " - (09) - " & UCase$(sKeywords(0)) & " (native) [+2]" & vbNewLine
'                                cLog.Append " > bType =" & sHexPrefix & Hex$(CByte(sKeywords(1))) & vbNewLine
'
'                            End If
                             
                        Case "wildbattle2"
                            
                            If Not PreProcessed Then
                                
                                If lOrgCount = 0 Then GoTo NoOrg
                                
                                lParamCount = 4&
                                
                                If lKeyCount < lParamCount + 1& Then
                                    GoTo TooLessParams
                                ElseIf lKeyCount > lParamCount + 1& Then
                                    GoTo TooMuchParams
                                End If
                                
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 10&
                                                                    
                            Else
                                
                                If Loc(iDestFile) <= lMaxFileSize - 10& Then
                                    
                                    Put #iDestFile, , CByte(&HB6)
                                    Put #iDestFile, , CWord(sKeywords(1))
                                    Put #iDestFile, , CByt(sKeywords(2))
                                    Put #iDestFile, , CWord(sKeywords(3))
                                    Put #iDestFile, , CByte(&H25)
                                    
                                    Select Case CByt(sKeywords(4))
                                        Case 0: Put #iDestFile, , CInt(&H137)
                                        Case 1: Put #iDestFile, , CInt(&H138)
                                        Case 2: Put #iDestFile, , CInt(&H139)
                                        Case 3: Put #iDestFile, , CInt(&H13A)
                                        Case 4: Put #iDestFile, , CInt(&H13B)
                                        Case 5: Put #iDestFile, , CInt(&H143)
                                        Case 6: Put #iDestFile, , CInt(&H156)
                                        Case Else: Put #iDestFile, , CInt(&H137)
                                    End Select
    
                                    Put #iDestFile, , CByte(&H27)
                                    
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                                
                                cLog.Append lLineNo & " - (--) - " & UCase$(sKeywords(0)) & " [+10]" & vbNewLine
                                cLog.Append " > iSpecies =" & sHexPrefix & Hex$(CInt(sKeywords(1))) & vbNewLine
                                cLog.Append " > bLevel =" & sHexPrefix & Hex$(CByte(sKeywords(2))) & vbNewLine
                                cLog.Append " > iHeldItem =" & sHexPrefix & Hex$(CInt(sKeywords(3))) & vbNewLine
                                cLog.Append " > bStyle =" & sHexPrefix & Hex$(CByte(sKeywords(4))) & vbNewLine

                            End If
                             
                        Case "else"
                            
                            If lOrgCount = 0 Then GoTo NoOrg
                            
                            lParamCount = 256&
                            
                            If lKeyCount < 2& Then
                                GoTo TooLessParams
                            End If
                            
                            Select Case sKeywords(1)
                                Case "call", "gosub"
                                    sKeywords(0) = "call"
                                    sKeywords(1) = sKeywords(2)
                                    lKeyCount = 2&
                                Case "goto", "jump"
                                    sKeywords(0) = "goto"
                                    sKeywords(1) = sKeywords(2)
                                    lKeyCount = 2&
                                Case Else
                                    sKeywords(0) = "call"
                            End Select
                            
                            GoTo Continue
                             
                        Case "giveitem3"
                            
                            If Not PreProcessed Then
                            
                                If lOrgCount = 0 Then GoTo NoOrg
                                
                                lParamCount = 1&
                                
                                If lKeyCount < lParamCount + 1& Then
                                    GoTo TooLessParams
                                ElseIf lKeyCount > lParamCount + 1& Then
                                    GoTo TooMuchParams
                                End If
                                
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 7&
                                                                    
                            Else
                                
                                If Loc(iDestFile) <= lMaxFileSize - 7& Then
                                    Put #iDestFile, , CByte(&H1A)
                                    Put #iDestFile, , CInt(&H8000)
                                    Put #iDestFile, , CWord(sKeywords(1))
                                    Put #iDestFile, , CInt(&H709)
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                                
                                cLog.Append lLineNo & " - (--) - " & UCase$(sKeywords(0)) & " [+7]" & vbNewLine
                                cLog.Append " > iItem =" & sHexPrefix & Hex$(CInt(sKeywords(1))) & vbNewLine
                            
                            End If
                         
                        Case "registernav"
                        
                            If sGameCode <> "BPE" And sGameCode <> String$(3, 0) Then
                                GoTo InvalidCommand
                            End If
                            
                            If Not PreProcessed Then
                                
                                If lOrgCount = 0 Then GoTo NoOrg
                                
                                lParamCount = 1&
                                
                                If lKeyCount < lParamCount + 1& Then
                                    GoTo TooLessParams
                                ElseIf lKeyCount > lParamCount + 1& Then
                                    GoTo TooMuchParams
                                End If
                                
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 7&
                                                                    
                            Else
                                
                                If Loc(iDestFile) <= lMaxFileSize - 7& Then
                                    Put #iDestFile, , CByte(&H1A)
                                    Put #iDestFile, , CInt(&H8000)
                                    Put #iDestFile, , CWord(sKeywords(1))
                                    Put #iDestFile, , CInt(&H809)
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                                
                                cLog.Append lLineNo & " - (--) - " & UCase$(sKeywords(0)) & " [+7]" & vbNewLine
                                cLog.Append " > iItem =" & sHexPrefix & Hex$(CInt(sKeywords(1))) & vbNewLine
                                
                            End If
                         
                        Case "cmdd3"
                                                               
                            If sGameCode <> "BPE" And sGameCode <> String$(3, 0) Then
                                GoTo InvalidCommand
                            End If
                                                              
                            If Not PreProcessed Then
                                
                                If lOrgCount = 0 Then GoTo NoOrg
                                
                                lParamCount = 1&
                                
                                If lKeyCount < lParamCount + 1& Then
                                    GoTo TooLessParams
                                ElseIf lKeyCount > lParamCount + 1& Then
                                    GoTo TooMuchParams
                                End If
                                
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 3&
                                                                    
                            Else
                                
                                If Loc(iDestFile) <= lMaxFileSize - 3& Then
                                    Put #iDestFile, , CByte(&HD3)
                                    Put #iDestFile, , CWord(sKeywords(1))
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                                
                                cLog.Append lLineNo & " - (D3) - CMDD3 [+3]" & vbNewLine
                                cLog.Append " > bByte =" & sHexPrefix & Hex$(CByte(sKeywords(1))) & vbNewLine
                                
                            End If
                             
                        Case "cmdd4"
                            
                            If sGameCode <> "BPE" And sGameCode <> String$(3, 0) Then
                                GoTo InvalidCommand
                            End If
                        
                            If Not PreProcessed Then
                            
                                If lOrgCount = 0 Then GoTo NoOrg
                                
                                lParamCount = 0&
                                
                                If lKeyCount < lParamCount + 1& Then
                                    GoTo TooLessParams
                                ElseIf lKeyCount > lParamCount + 1& Then
                                    GoTo TooMuchParams
                                End If
                                
                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 1&
                                
                            Else
                                
                                If Loc(iDestFile) <= lMaxFileSize - 1& Then
                                    Put #iDestFile, , CByte(&HD4)
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                                
                                cLog.Append lLineNo & " - (D4) - CMDD4 [+1]" & vbNewLine
                                
                            End If
               
                        Case Else
                            
                            If lOrgCount = 0 Then
                                GoTo NoOrg
                            End If
                            
                            sErrorDescr = LoadResString(13014) & Replace(sKeywords(0), "&H", "0x") & LoadResString(13015) & lLineNo & "."
                    
                            If LenB(sScriptFile) <> 0 Then
                                sErrorDescr = sErrorDescr & vbNewLine & LoadResString(13033) & ": """ & sScriptFile & """."
                            End If
                            
                            GoTo ErrorMessage
                         
                    End Select
                               
                End If
                 
            Else
                 
                Select Case sKeywords(0)
                    
                    '----- DIRECTIVES -----
                    Case "#org", "#seek"
                         
                        If Not PreProcessed Then
                         
                            lOrgCount = lOrgCount + 1&
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            If AscW(sKeywords(1)) = ascAt Then
                            
                                If Len(sKeywords(1)) > 1& Then
                                
                                    If lDynamicStart > -1& Then
                                        
                                        If FirstOrg = False Then

                                            If lCurrentOrg = UBound(DynamicOffsets) Then
                                                ReDim Preserve DynamicOffsets(lCurrentOrg + 9&) As Define
                                                ReDim Preserve lTotalBytes(lCurrentOrg + 9&) As Long
                                            End If
                                            
                                            lCurrentOrg = lCurrentOrg + 1&
                                            
                                        Else
                                            FirstOrg = False
                                        End If
                                        
                                        lDynamicCount = lDynamicCount + 1&
                                        DynamicOffsets(lCurrentOrg).Symbol = sKeywords(1)
                                        colTest.Add 0, sKeywords(1)
                                        
                                    Else
                                        GoTo NoDynamic
                                    End If
                                    
                                Else
                                    MissingDynamic = True
                                    Err.Raise TypeMismatch
                                End If
                                
                            End If
                                                                        
                        Else
                            
                            If AscW(sKeywords(1)) = ascAt Then
                                
                                If Not FirstOrg Then
                                    lCurrentOrg = lCurrentOrg + 1&
                                Else
                                    FirstOrg = False
                                End If
                                
                                sKeywords(1) = DynamicOffsets(lCurrentOrg).Value
                                
                            End If
                            
                            If CLng(sKeywords(1)) <= lMaxFileSize Then
                                Seek #iDestFile, CLng(sKeywords(1)) + 1&
                            Else
                                Err.Raise BadRecordNumber
                            End If
                            
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > lNewOffset =" & sHexPrefix & Hex$(CLng(sKeywords(1))) & vbNewLine

                        End If
                        
                    Case "#raw", "#binary", "#put"
                        
                        If Not PreProcessed Then
                        
                            If lOrgCount = 0 Then GoTo NoOrg
                            
                            lParamCount = 256&
                            
                            If lKeyCount < 2& Then
                                GoTo TooLessParams
                            End If
                        
                        Else
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                        End If
                        
                        bTemp = 0
                        
                        For i = 1 To lKeyCount - 1
                            
                            Select Case sKeywords(i)
                            
                                Case "word", "i", "int", "integer"
                                    bTemp = 1
                                Case "pointer", "p", "ptr"
                                    bTemp = 3
                                Case "byte", "b", "char"
                                    bTemp = 0
                                Case "dword", "l", "long"
                                    bTemp = 2
                                
                                Case Else  'It's a value
                                    
                                    If Not PreProcessed Then
                                    
                                        Select Case bTemp
                                            Case 0
                                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 1&
                                            Case 1
                                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 2&
                                            Case 3, 2
                                                lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + 4&
                                        End Select
                                        
                                    Else
                                    
'                                        If InStrB(sKeywords(i), "&H") <> 1 Then
'                                            If IsHex(sKeywords(i)) Then
'                                                sKeywords(i) = "&H" & sKeywords(i)
'                                            End If
'                                        End If
                                        
                                        Select Case bTemp
                                        
                                            Case 0
                                            
                                                If Loc(iDestFile) <= lMaxFileSize - 1& Then
                                                    Put #iDestFile, , CByt(sKeywords(i))
                                                    cLog.Append " > bOut =" & sHexPrefix & Hex$(CByte(sKeywords(i))) & vbNewLine
                                                Else
                                                    Err.Raise BadRecordNumber
                                                End If
                                                
                                            Case 1
                                            
                                                If Loc(iDestFile) <= lMaxFileSize - 2& Then
                                                    Put #iDestFile, , CWord(sKeywords(i))
                                                    cLog.Append " > iOut =" & sHexPrefix & Hex$(CInt(sKeywords(i))) & vbNewLine
                                                Else
                                                    Err.Raise BadRecordNumber
                                                End If
                                                
                                            Case 3
                                            
                                                If Loc(iDestFile) <= lMaxFileSize - 4& Then
                                                    Put #iDestFile, , CPtr(sKeywords(i))
                                                    cLog.Append " > pOut =" & sHexPrefix & Hex$(CPtr(sKeywords(i))) & vbNewLine
                                                Else
                                                    Err.Raise BadRecordNumber
                                                End If
                                                
                                            Case 2
                                            
                                                If Loc(iDestFile) <= lMaxFileSize - 4& Then
                                                    Put #iDestFile, , CLong(sKeywords(i))
                                                    cLog.Append " > lOut =" & sHexPrefix & Hex$(CLng(sKeywords(i))) & vbNewLine
                                                Else
                                                    Err.Raise BadRecordNumber
                                                End If
                                                
                                        End Select
                                        
                                    End If
                                    
                            End Select
                            
                        Next i
                        
                    Case "#dynamic"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            DoReplace sKeywords(1), "0x", "&H"
                            
                            If CLng(sKeywords(1)) <= lMaxFileSize Then
                                lDynamicStart = CLng(sKeywords(1))
                            Else
                                If IsDebugging Then
                                    lMaxFileSize = lMaxFileSize + CLng(sKeywords(1))
                                    lDynamicStart = CLng(sKeywords(1))
                                Else
                                    Err.Raise BadRecordNumber
                                End If
                            End If
                        
                        Else

                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > lDynamicStart =" & sHexPrefix & Hex$(CLng(sKeywords(1))) & vbNewLine
                            
                        End If
                        
                    Case "#include"
                        
                        If Not PreProcessed Then
                            
                            If lKeyCount >= 2& Then
                                For i = 1 To lKeyCount - 1&
                                    HeaderProcess sKeywords(i)
                                Next i
                            Else
                                lParamCount = 256&
                                GoTo TooLessParams
                            End If
                            
                        Else
                        
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            
                            For i = 1 To lKeyCount - 1
                                cLog.Append " > sFile = " & sKeywords(i) & vbNewLine
                            Next i
                                                            
                        End If
                        
                    Case "#alias"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 2&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            AddAlias sKeywords(1), sKeywords(2)
                            cTemp.Insert cTemp.Length - 2& - Len(sKeywords(2)) - 1& - Len(sKeywords(1)) + 1&, "¦"
                            cTemp.Insert cTemp.Length - 2& - Len(sKeywords(2)) + 1&, "¦"
                            
                        Else

                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > sSymbol = " & Replace(sKeywords(1), "¦", vbNullString) & vbNewLine
                            cLog.Append " > sAlias = " & Replace(sKeywords(2), "¦", vbNullString) & vbNewLine
                            
                        End If
                        
                    Case "#define", "#const"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 2&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            AddDefine sKeywords(1), sKeywords(2)
                            cTemp.Insert cTemp.Length - 2& - Len(sKeywords(2)) - 1& - Len(sKeywords(1)) + 1&, "¦"
                            
                        Else
                            
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > sSymbol = " & Replace(UCase$(sKeywords(1)), "¦", vbNullString) & vbNewLine
                            cLog.Append " > sValue = " & Replace(UCase$(sKeywords(2)), "&H", "0x") & vbNewLine
                            
                        End If
                                                        
                    Case "#reserve"
                        
                        If Not PreProcessed Then
                            
                            If lOrgCount = 0 Then GoTo NoOrg
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            DoReplace sKeywords(1), "0x", "&H"
                            lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + CLong(sKeywords(1))
                            
                        Else
                            
                            If CLng(sKeywords(1)) <= lMaxFileSize - Loc(iDestFile) Then
                                
                                If CLng(sKeywords(1)) > 0 Then
                                
                                    ReDim bTempArray(CLng(sKeywords(1)) - 1) As Byte
                                    RtlFillMemory bTempArray(0), UBound(bTempArray) + 1&, &H1
                                    
                                    Put #iDestFile, , bTempArray
                                    Erase bTempArray
                                
                                End If
                                
                            Else
                                Err.Raise BadRecordNumber
                            End If
                            
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > lByteAmount =" & sHexPrefix & Hex$(sKeywords(1)) & vbNewLine
                            
                        End If
                        
                    Case "#clean"
                                               
                        If Not PreProcessed Then
                            
                            lParamCount = 0&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            Debug.Assert App.hInstance
                            
                            If (Not Not JunkData) <> 0 Then
                                
                                For i = LBound(JunkData) To UBound(JunkData)
                                    
                                    ReDim bTempArray(JunkData(i).Length - 1&) As Byte
                                    
                                    If bFreeSpace <> 0 Then
                                        RtlFillMemory bTempArray(0), UBound(bTempArray) + 1&, bFreeSpace
                                    End If
                                    
                                    Put #iDestFile, JunkData(i).Offset + 1&, bTempArray
                                    
                                Next i
                                
                                Erase JunkData
                                Erase bTempArray
                                
                            End If
                        
                        Else
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                        End If
                                               
                    Case "#erase"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 2&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            DoReplace sKeywords(1), "0x", "&H"
                            DoReplace sKeywords(2), "0x", "&H"

                            If CLng(sKeywords(1)) <= lMaxFileSize And CLng(sKeywords(2)) <= lMaxFileSize - CLng(sKeywords(1)) Then
                                
                                If CLng(sKeywords(2)) > 0 Then
                                    
                                    ReDim bTempArray(CLng(sKeywords(2)) - 1&) As Byte
                                    
                                    If bFreeSpace <> 0 Then
                                        RtlFillMemory bTempArray(0), UBound(bTempArray) + 1&, bFreeSpace
                                    End If
                                    
                                    Put #iDestFile, CLng(sKeywords(1)) + 1&, bTempArray
                                    Erase bTempArray
                                
                                End If
                                
                            Else
                                Err.Raise BadRecordNumber
                            End If
                            
                        Else
                            
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > lStartOffset =" & sHexPrefix & Hex$(sKeywords(1)) & vbNewLine
                            cLog.Append " > lLength =" & sHexPrefix & Hex$(sKeywords(2)) & vbNewLine
                            
                        End If
                        
                    Case "#eraserange"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 2&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            DoReplace sKeywords(1), "0x", "&H"
                            DoReplace sKeywords(2), "0x", "&H"
                            
                            If CLng(sKeywords(1)) <= lMaxFileSize And CLng(sKeywords(2)) <= lMaxFileSize Then
                                If CLng(sKeywords(1)) <= CLng(sKeywords(2)) Then
                                
                                    If (CLng(sKeywords(2)) - CLng(sKeywords(1))) > 0 Then
                                
                                        ReDim bTempArray((CLng(sKeywords(2)) - CLng(sKeywords(1))) - 1&) As Byte
                                        
                                        If bFreeSpace <> 0 Then
                                            RtlFillMemory bTempArray(0), UBound(bTempArray) + 1&, bFreeSpace
                                        End If
                                        
                                        Put #iDestFile, CLng(sKeywords(1)) + 1&, bTempArray
                                        Erase bTempArray
                                    
                                    End If
                                    
                                Else
                                    
                                    If (CLng(sKeywords(1)) - CLng(sKeywords(2))) > 0 Then
                                    
                                        ReDim bTempArray((CLng(sKeywords(1)) - CLng(sKeywords(2))) - 1) As Byte
                                        
                                        If bFreeSpace <> 0 Then
                                            RtlFillMemory bTempArray(0), UBound(bTempArray) + 1&, bFreeSpace
                                        End If
                                        
                                        Put #iDestFile, CLng(sKeywords(2)) + 1, bTempArray
                                        Erase bTempArray
                                        
                                    End If
                                    
                                End If
                            Else
                                Err.Raise BadRecordNumber
                            End If
                            
                        Else
                        
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > lStartOffset =" & sHexPrefix & Hex$(sKeywords(1)) & vbNewLine
                            cLog.Append " > lEndOffset =" & sHexPrefix & Hex$(sKeywords(2)) & vbNewLine
                            
                        End If
                        
                    Case "#removeall"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            DoReplace sKeywords(1), "0x", "&H"
                            
                            If CLng(sKeywords(1)) <= lMaxFileSize Then
                                Decompile sFileName, sKeywords(1), 2
                            Else
                                Err.Raise BadRecordNumber
                            End If
                            
                        Else

                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > lScriptOffset =" & sHexPrefix & Hex$(sKeywords(1)) & vbNewLine
                            
                        End If
                        
                    Case "#remove"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            DoReplace sKeywords(1), "0x", "&H"
                            
                            If CLng(sKeywords(1)) <= lMaxFileSize Then
                                
                                lTemp = Decompile(sFileName, sKeywords(1), 1)
                                
                                If lTemp > 0 Then
                                
                                    ReDim bTempArray((lTemp - CLng(sKeywords(1))) - 1&) As Byte
                                    
                                    If bFreeSpace <> 0 Then
                                        RtlFillMemory bTempArray(0), UBound(bTempArray) + 1&, bFreeSpace
                                    End If
                                    
                                    Put #iDestFile, CLng(sKeywords(1)) + 1&, bTempArray
                                    Erase bTempArray
                                    
                                End If
                                
                            Else
                                Err.Raise BadRecordNumber
                            End If
                            
                        Else

                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > lScriptOffset =" & sHexPrefix & Hex$(sKeywords(1)) & vbNewLine
                            
                        End If
                        
                    Case "#removestring"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            DoReplace sKeywords(1), "0x", "&H"
                            
                            If CLng(sKeywords(1)) <= lMaxFileSize Then
                                
                                ReDim bTempArray(999) As Byte
                                Get #iDestFile, CLng(sKeywords(1)) + 1&, bTempArray
                                
                                ReDim bTempArray2(0) As Byte: bTempArray2(0) = bFreeSpace
                                lTemp = InStrB(bTempArray, bTempArray2)
                                
                                If lTemp > 0& Then
                                    If bTempArray(lTemp) = 0 Then
                                        lTemp = lTemp + 1&
                                    End If
                                End If
                                
                                If lTemp > 1& Then
                                
                                    ReDim bTempArray2(lTemp - 1&) As Byte
                                                        
                                    If bFreeSpace <> 0 Then
                                        RtlFillMemory bTempArray2(0), UBound(bTempArray2) + 1&, bFreeSpace
                                    End If
                                    
                                    Put #iDestFile, CLng(sKeywords(1)) + 1&, bTempArray2
                                    Erase bTempArray2
                                    
                                End If
                                
                            Else
                                Err.Raise BadRecordNumber
                            End If
                            
                        Else

                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > lOffset =" & sHexPrefix & Hex$(sKeywords(1)) & vbNewLine
                            
                        End If
                        
                    Case "#removemove"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            DoReplace sKeywords(1), "0x", "&H"
                            
                            If CLng(sKeywords(1)) <= lMaxFileSize Then
                                
                                ReDim bTempArray(511) As Byte
                                Get #iDestFile, CLng(sKeywords(1)) + 1&, bTempArray
                                
                                ReDim bTempArray2(0) As Byte: bTempArray2(0) = &HFE
                                lTemp = InStrB(bTempArray, bTempArray2)
                                
                                If lTemp > 0& Then
                                
                                    ReDim bTempArray2(lTemp - 1&) As Byte
                                                        
                                    If bFreeSpace <> 0 Then
                                        RtlFillMemory bTempArray2(0), UBound(bTempArray2) + 1&, bFreeSpace
                                    End If
                                    
                                    Put #iDestFile, CLng(sKeywords(1)) + 1&, bTempArray2
                                    Erase bTempArray2
                                    
                                End If
                                
                            Else
                                Err.Raise BadRecordNumber
                            End If
                            
                        Else

                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > lOffset =" & sHexPrefix & Hex$(sKeywords(1)) & vbNewLine
                            
                        End If
                        
                    Case "#removemart"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            DoReplace sKeywords(1), "0x", "&H"
                            
                            If CLng(sKeywords(1)) <= lMaxFileSize Then
                                
                                ReDim bTempArray(199) As Byte
                                Get #iDestFile, CLng(sKeywords(1)) + 1&, bTempArray
                                
                                ReDim bTempArray2(1) As Byte
                                lTemp = InStrB(bTempArray, bTempArray2)
                                
                                If lTemp > 0& Then
                                    
                                    If lTemp > 1 Then
                                        lTemp = lTemp + 2&
                                    Else
                                        lTemp = 2&
                                    End If
                                    
                                    ReDim bTempArray2(lTemp - 1&) As Byte
                                                        
                                    If bFreeSpace <> 0 Then
                                        RtlFillMemory bTempArray2(0), UBound(bTempArray2) + 1&, bFreeSpace
                                    End If
                                    
                                    Put #iDestFile, CLng(sKeywords(1)) + 1&, bTempArray2
                                    Erase bTempArray2
                                    
                                End If
                                
                            Else
                                Err.Raise BadRecordNumber
                            End If
                            
                        Else

                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > lOffset =" & sHexPrefix & Hex$(sKeywords(1)) & vbNewLine
                            
                        End If
                        
                    Case "#braille"

                        If Not PreProcessed Then
                        
                            If lOrgCount = 0 Then
                                GoTo NoOrg
                            End If
                            
                            TextData(lTextCount).sText = UCase$(MidB$(sRawInput, 19&))
                                
                            lTemp = Asc2BrailleLen(TextData(lTextCount).sText) + 1&
                            ReDim TextData(lTextCount).bArray(lTemp - 1) As Byte
                            
                            Asc2Braille TextData(lTextCount).sText & "\X", TextData(lTextCount).bArray
                            
                            If lTextCount = UBound(TextData) Then
                                ReDim Preserve TextData(lTextCount + 19&) As ByteString
                            End If
                            
                            lTextCount = lTextCount + 1&
                            lTotalBytes(lCurrentOrg) = lTotalBytes(lCurrentOrg) + lTemp + 1&
                            
                        Else
                            
                            If Loc(iDestFile) <= lMaxFileSize - (UBound(TextData(lTextCount).bArray) + 1&) Then
                                Put #iDestFile, , TextData(lTextCount).bArray
                            Else
                                Err.Raise BadRecordNumber
                            End If
                            
                            If LenB(DynamicOffsets(lCurrentOrg).Symbol) <> 0 Then
                                Put #iDestFile, , CByte(&H0)
                            End If
                            
                            cLog.Append lLineNo & " - BRAILLE TEXT [+" & (UBound(TextData(lTextCount).bArray) + 1&) & "]" & vbNewLine
                            cLog.Append " > sText = """ & TextData(lTextCount).sText & """" & vbNewLine
                            
                            lTextCount = lTextCount + 1&
                            
                        End If
                        
                    Case "#autobank"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                        Else
                            
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            
                            Select Case sKeywords(1)
                                Case "off", "0"
                                    AutoBank = False
                                    cLog.Append LoadResString(13006) & vbNewLine
                                Case "on", "1"
                                    AutoBank = True
                                    cLog.Append LoadResString(13007) & vbNewLine
                                Case Else
                                    AutoBank = True
                                    cLog.Append LoadResString(13008) & vbNewLine
                                    cLog.Append LoadResString(13009) & vbNewLine
                            End Select
                            
                        End If
                        
                    Case "#freespace"

                        If Not PreProcessed Then
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            DoReplace sKeywords(1), "0x", "&H"
                            
                            If IsDebugging = False Then
                                If CByte(sKeywords(1)) = 0 Then
                                    bFreeSpace = 0
                                Else
                                    bFreeSpace = &HFF
                                End If
                            End If
                            
                        Else

                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > bFreeSpace =" & sHexPrefix & Hex$(sKeywords(1)) & vbNewLine
                            
                        End If
                                                        
                    Case "#definelist", "#constlist"
                         
                        If Not PreProcessed Then
                            
                            lParamCount = 0&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                        Else
                            
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            
                            If Defines.Count > 0 Then
                                            
                                cLog.Append String$(37, "-") & vbNewLine
                            
                                For i = 1 To Defines.Count
                                    cLog.Append "#" & RightB$(String$(Len(CStr(Defines.Count - 1&)), "0") & i, Len(CStr(Defines.Count - 1&)) * 2&) & " - [" & ColItemKey(i, Defines) & "] > " & Defines.Item(i) & vbNewLine
                                Next i
                                
                                cLog.Append String$(37, "-") & vbNewLine
                            
                            End If
                            
                        End If
                                                        
                    Case "#break", "#stop"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 0&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                        Else
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                        End If
                        
                        Exit Do
                        
                    Case "#undefine", "#deconst"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            RemoveDefine (sKeywords(1))
                            cTemp.Insert cTemp.Length - 2 - Len(sKeywords(1)) + 1, "¦"
                            
                        Else
                            
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > sSymbol = " & Replace(UCase$(sKeywords(1)), "¦", vbNullString) & vbNewLine
                            
                        End If
                        
                    Case "#undefineall", "#deconstall"
                    
                        If Not PreProcessed Then
                            
                            lParamCount = 0&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            EraseCol Defines
                            
                        Else
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                        End If
                        
                    Case "#unalias"
                        
                        If Not PreProcessed Then
                            
                            lParamCount = 1&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            RemoveAlias (sKeywords(1))
                            cTemp.Insert cTemp.Length - 2& - Len(sKeywords(1)) + 1&, "¦"
                        
                        Else
                        
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                            cLog.Append " > sAlias = " & Replace(sKeywords(1), "¦", vbNullString) & vbNewLine
                            
                        End If
                        
                    Case "#unaliasall"
                    
                        If Not PreProcessed Then
                            
                            lParamCount = 0&
                            
                            If lKeyCount < lParamCount + 1& Then
                                GoTo TooLessParams
                            ElseIf lKeyCount > lParamCount + 1& Then
                                GoTo TooMuchParams
                            End If
                            
                            EraseCol Aliases
                            
                        Else
                            cLog.Append lLineNo & " - " & UCase$(sKeywords(0)) & vbNewLine
                        End If
                        
                    Case Else
                    
                        sErrorDescr = LoadResString(13014) & Replace(sKeywords(0), "&H", "0x") & LoadResString(13015) & lLineNo & "."
                        
                        If LenB(sScriptFile) <> 0 Then
                            sErrorDescr = sErrorDescr & vbNewLine & LoadResString(13033) & ": """ & sScriptFile & """."
                        End If
                        
                        GoSub ErrorMessage
                         
                End Select

            End If
            
            If Not PreProcessed Then
                If AlreadyAdded Then
                    sRawInput = sTemp
                    AlreadyAdded = False
                    GoTo Continue
                End If
            End If
            
        Loop
        
        Close #iFileNum
          
        If Not PreProcessed Then
            
            lLineNo = 0
            
            sTemp = cTemp.ToString
          
            If lDynamicCount > 0 Then
          
                ReDim bBuffer(lChunkSize - 1&) As Byte
                        
                For i = 0& To lDynamicCount - 1&
                        
                    ReDim bSearch(lTotalBytes(i)) As Byte
                    
                    If bFreeSpace = &HFF Then
                        RtlFillMemory bSearch(0), lTotalBytes(i) + 1&, &HFF
                    End If
                    
                    For j = 0& To (lMaxFileSize - lDynamicStart) \ lChunkSize
                        
                        Get #iDestFile, lDynamicStart + 1&, bBuffer
                        lFoundOffset = InStrB(bBuffer, bSearch)
                        
                        If lFoundOffset <> 0 Then
                            Exit For
                        End If
                        
                        lDynamicStart = lDynamicStart + lChunkSize
                        
                    Next j
                                          
                    'DynamicOffsets(i).Value = "&H" & Hex$((lFoundOffset - 1&) + lDynamicStart)
                    DynamicOffsets(i).Value = (lFoundOffset - 1&) + lDynamicStart
                    
                    If CLng(DynamicOffsets(i).Value) = 0 Then
                        If lFoundOffset <> 1 Then
                            GoTo NotEnoughSpace
                        End If
                    ElseIf CLng(DynamicOffsets(i).Value) > lMaxFileSize - 2& Then
                        GoTo NotEnoughSpace
                    End If
                    
                    lDynamicStart = lDynamicStart + lFoundOffset + lTotalBytes(i)
                                                
                Next i
                  
                Erase bBuffer
                Erase bSearch
                Set colTest = Nothing
                
                ReDim Preserve DynamicOffsets(lDynamicCount - 1&) As Define
                
                DynamicOffsets2 = DynamicOffsets
                TriQuickSortDefine DynamicOffsets2
'
'                For i = 0 To lDynamicCount - 1
'                    DoReplace sTemp, DynamicOffsets2(i).Symbol, "&H" & Hex$(DynamicOffsets2(i).Value)
'                Next i
                
                If IsDebugging = False Then

                    If lTotalBytes(0) > 0 Then
                    
                        lTemp = 1&
                        
                        For i = 1 To lDynamicCount - 1&
                            If lTotalBytes(i) <> 0 Then
                                lTemp = lTemp + 1&
                            Else
                                Exit For
                            End If
                        Next i

                        ReDim JunkData(lTemp - 1&) As tJunk
                        
                        For i = LBound(JunkData) To UBound(JunkData)
                            JunkData(i).Offset = DynamicOffsets(i).Value
                            JunkData(i).Length = lTotalBytes(i)
                        Next i
                        
                    End If
                
                End If
                
                'Erase DynamicOffsets2
                Erase lTotalBytes
                      
            End If
            
            If Defines.Count > 0 Then
                                  
                ReDim Defines2(Defines.Count + Aliases.Count - 1&) As Define
                
                For i = 1& To Defines.Count
                    Defines2(i - 1&).Symbol = ColItemKey(i, Defines)
                    Defines2(i - 1&).Value = Defines.Item(i)
                Next i
                
                If Aliases.Count > 0 Then
                    For i = 1& To Aliases.Count
                        Defines2(i - 1& + Defines.Count).Symbol = Aliases.Item(i)
                        Defines2(i - 1& + Defines.Count).Value = ColItemKey(i, Aliases)
                    Next i
                End If
                    
                TriQuickSortDefine Defines2, SortDescending
                
                For i = LBound(Defines2) To UBound(Defines2)
                    DoReplace sTemp, Defines2(i).Symbol, Defines2(i).Value
                Next i
                    
                Erase Defines2
            
            End If
            
            If Aliases.Count > 0 Then
                                  
                If Defines.Count = 0 Then
                                                  
                    ReDim Defines2(Aliases.Count - 1&) As Define

                    For i = 1 To Aliases.Count
                        Defines2(i - 1&).Symbol = Aliases.Item(i)
                        Defines2(i - 1&).Value = ColItemKey(i, Aliases)
                    Next i

                    TriQuickSortDefine Defines2, SortDescending
                    
                    For i = LBound(Defines2) To UBound(Defines2)
                        DoReplace sTemp, Defines2(i).Symbol, Defines2(i).Value
                    Next i

                    Erase Defines2

                End If

            End If
            
            lTextCount = 0
            lCurrentOrg = 0
            FirstOrg = True
            PreProcessed = True
            SplitB sTemp, sArray, vbNewLine, lLineCount
            
            GoTo Begin
            
        Else
            
            lTemp = EndTiming
            SetStatusText LoadResString(13026) & Format$(lTemp / 1000, "0.000") & LoadResString(13025)
            
            If NoLog = False Then
                
                If lDynamicCount > 0 Then
                  
                    If Batch = False Then
                        frmOutput.lstDynamics.Clear
                        frmOutput.lstOffsets.Clear
                    Else
                        AddItem frmOutput, frmOutput.lstDynamics, "[" & GetFileName(sScriptFile) & "]"
                        AddItem frmOutput, frmOutput.lstOffsets, vbSpace
                    End If
                  
                    cLog.Append String$(37, "-") & vbNewLine
                  
                    LockUpdate frmOutput.lstDynamics.hWnd
                    LockUpdate frmOutput.lstOffsets.hWnd
                  
                    For i = 0 To lDynamicCount - 1&
                        
                        AddItem frmOutput, frmOutput.lstDynamics, DynamicOffsets(i).Symbol
                        AddItem frmOutput, frmOutput.lstOffsets, Hex$(DynamicOffsets(i).Value)
                        
                        cLog.Append "DYNAMIC_OFFSET " & i + 1 & vbNewLine
                        cLog.Append " > sLabel = " & DynamicOffsets(i).Symbol & vbNewLine
                        cLog.Append " > lOffset =" & sHexPrefix & Hex$(DynamicOffsets(i).Value) & vbNewLine
                        
                    Next i
                    
                    UnlockUpdate frmOutput.lstDynamics.hWnd
                    UnlockUpdate frmOutput.lstOffsets.hWnd
                    
                    If Batch = False Then
                        frmOutput.lstDynamics.ListIndex = 0
                        frmOutput.lstOffsets.ListIndex = 0
                    Else
                        frmOutput.lstDynamics.ListIndex = 1
                        frmOutput.lstOffsets.ListIndex = 1
                    End If
                    
                    frmOutput.Height = 6240
                  
              Else
                  frmOutput.Height = 4515
              End If
              
              If lLineNo > 0 Then
                cLog.Append String$(37, "-") & vbNewLine
              End If
              
              cLog.Append LoadResString(13022) & vbNewLine
              cLog.Append LoadResString(13023) & vbNewLine
              cLog.Append LoadResString(13024) & Format$(lTemp / 1000, "0.000") & LoadResString(13025)
              
              If Batch = True Then
                  If LenB(frmOutput.txtOutput.text) <> 0 Then
                      cLog.Append vbNewLine & vbNewLine
                      cLog.Append frmOutput.txtOutput.text
                  End If
              End If
              
              SendMessageW frmOutput.txtOutput.hWnd, WM_SETTEXT, 0&, ByVal cLog.Pointer
              Show2 frmOutput, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
              
            Else
                
                If IsDebugging = False Then
                    If lDynamicCount > 0 Then
                        SafeClipboardSet Hex$(DynamicOffsets(0).Value)
                    End If
                End If
                
            End If
            
            Set cLog = Nothing
            
        End If
        
    Close #iDestFile
    ClearData
    
    Exit Function
      
StdErr:
    
    If lLineNo > 0 Then
        sErrorDescr = LoadResString(13016) & Err.Number & " """ & Err.Description & """" & LoadResString(13017) & lLineNo & "." & vbNewLine
    Else
        sErrorDescr = LoadResString(13016) & Err.Number & " """ & Err.Description & """." & vbNewLine
    End If
      
    If LenB(sScriptFile) <> 0 Then
        sErrorDescr = sErrorDescr & LoadResString(13033) & ": """ & sScriptFile & """." & vbNewLine
    End If
    
    Select Case Err.Number
        
        Case Overflow
            sErrorDescr = sErrorDescr & LoadResString(13020) & vbNewLine
            
        Case TypeMismatch
            
            If MissingDynamic Then
                sErrorDescr = sErrorDescr & LoadResString(13019) & vbNewLine
                MissingDynamic = False
            ElseIf MissingDefine Then
                sErrorDescr = sErrorDescr & LoadResString(13048) & vbNewLine
                MissingDefine = False
            End If
            
        Case FileAccessErr
            
            sErrorDescr = LoadResString(13016) & Err.Number & " """ & Err.Description & """" & "."
            
            If TriedUnlocking = False Then
                
                TriedUnlocking = True
                
                If UnLockFile(sFileName) Then
                    Close #iDestFile, #iFileNum
                    Process sFileName, sScriptFile, Batch
                    Exit Function
                End If

            End If
        
        Case DuplicateKey
            GoTo FoundDuplicate
            
    End Select
    
    GoTo ErrorMessage
    
NoOrg:
    
    sErrorDescr = LoadResString(13028)
    
    If LenB(sScriptFile) <> 0 Then
        sErrorDescr = sErrorDescr & vbNewLine & LoadResString(13033) & ": """ & sScriptFile & """."
    End If
    
    GoTo ErrorMessage
    
TooLessParams:
    
    If lParamCount <> 256 Then
        sErrorDescr = LoadResString(13046) & LoadResString(13017) & lLineNo & ". " & LoadResString(13047) & lParamCount & "."
    Else
        sErrorDescr = LoadResString(13046) & LoadResString(13017) & lLineNo & "."
    End If
    
    If LenB(sScriptFile) <> 0 Then
        sErrorDescr = sErrorDescr & vbNewLine & LoadResString(13033) & ": """ & sScriptFile & """."
    End If
    
    GoTo ErrorMessage
    
TooMuchParams:

    sErrorDescr = LoadResString(13045) & LoadResString(13017) & lLineNo & ". " & LoadResString(13047) & lParamCount & "."
    
    If LenB(sScriptFile) <> 0 Then
        sErrorDescr = sErrorDescr & vbNewLine & LoadResString(13033) & ": """ & sScriptFile & """."
    End If
    
    GoTo ErrorMessage
      
InvalidCommand:
    
    sErrorDescr = LoadResString(13044) & Replace(sKeywords(0), "&H", "0x") & LoadResString(13015) & lLineNo & "."
    
    If LenB(sScriptFile) <> 0 Then
        sErrorDescr = sErrorDescr & vbNewLine & LoadResString(13033) & ": """ & sScriptFile & """."
    End If
    
    GoTo ErrorMessage

NoDynamic:
    
    sErrorDescr = LoadResString(13027)
    
    If LenB(sScriptFile) <> 0 Then
        sErrorDescr = sErrorDescr & vbNewLine & LoadResString(13033) & ": """ & sScriptFile & """."
    End If
    
    GoTo ErrorMessage
   
FoundDuplicate:
    
    sErrorDescr = LoadResString(13029) & lLineNo & "."
    
    If LenB(sScriptFile) <> 0 Then
        sErrorDescr = sErrorDescr & vbNewLine & LoadResString(13033) & ": """ & sScriptFile & """."
    End If
    
    GoTo ErrorMessage
   
NotEnoughSpace:
    
    sErrorDescr = LoadResString(13042)
    
    If LenB(sScriptFile) <> 0 Then
        sErrorDescr = sErrorDescr & vbNewLine & LoadResString(13033) & ": """ & sScriptFile & """."
    End If
    
    GoTo ErrorMessage
    
ErrorMessage:
    
    MsgBox sErrorDescr, vbExclamation
    
    If Batch = False Then
        GotoLine (lLineNo)
    End If
    
Terminate:
    Process = False
    Close #iDestFile, #iFileNum
    ClearData
    DeleteFile sTempPath & sTempFile
    DeleteFile sTempPath & sTempLog
    frmMain.WelcomeText
   
End Function

Private Function IsPtr(lOffset As Long) As Boolean
    If (lOffset And &HFF000000) >= &H8000000 Then
        If (lOffset And &HFF000000) <= &H9000000 Then
            IsPtr = True
        End If
    End If
End Function

Private Sub AddSnip(lOffset As Long)
    
    On Error GoTo AlreadyIn
    
    If IsPtr(lOffset) Then
        lOffset = (lOffset And &HFFFFFFF) - &H8000000
        Snips.Add lOffset, CStr(lOffset)
    End If
    
AlreadyIn:
End Sub

Private Sub AddString(lOffset As Long)
        
    On Error GoTo AlreadyIn
    
    If IsPtr(lOffset) Then
        lOffset = (lOffset And &HFFFFFFF) - &H8000000
        Strings.Add lOffset, CStr(lOffset)
    End If
    
AlreadyIn:
End Sub

Private Sub AddMove(lOffset As Long)
    
    On Error GoTo AlreadyIn
    
    If IsPtr(lOffset) Then
        lOffset = (lOffset And &HFFFFFFF) - &H8000000
        Moves.Add lOffset, CStr(lOffset)
    End If
    
AlreadyIn:
End Sub

Private Sub AddMart(lOffset As Long)
    
    On Error GoTo AlreadyIn
    
    If IsPtr(lOffset) Then
        lOffset = (lOffset And &HFFFFFFF) - &H8000000
        Marts.Add lOffset, CStr(lOffset)
    End If
    
AlreadyIn:
End Sub

Private Sub AddBraille(lOffset As Long)
    
    On Error GoTo AlreadyIn
    
    'If IsPtr(lOffset) Then
        lOffset = (lOffset And &HFFFFFFF) - &H8000000
        Brailles.Add lOffset, CStr(lOffset)
    'End If
    
AlreadyIn:
End Sub

Private Function GetHeaderDefine(sHeader As String, iIndex As Integer) As String
Dim iFileNum As Integer
Dim sInput As String
Dim sArray() As String
Dim i As Long

    iFileNum = FreeFile
    
    If FileExists(sHeader) Then
    
        Open sHeader For Input As #iFileNum
        
            On Error GoTo SomethingWrong
            
            For i = 0 To iIndex
                Line Input #iFileNum, sInput
            Next i
            
        Close #iFileNum
        
        SplitB sInput, sArray, vbSpace
        
        If UBound(sArray) > 0 Then
            GetHeaderDefine = vbSpace & sArray(1)
        Else
            GoTo SomethingWrong
        End If
    
    Else
    
SomethingWrong:
        GetHeaderDefine = sHexPrefix & Hex$(iIndex)
    End If

End Function

Public Function Decompile(sFileName As String, ByVal lOffset As Long, Optional ByRef Remove As Integer = 0) As Long
Dim tmpByte1 As Byte
Dim tmpByte2 As Byte
Dim tmpByte3 As Byte
Dim tmpByte4 As Byte
Dim tmpInt1 As Integer
Dim tmpInt2 As Integer
Dim tmpInt3 As Integer
Dim tmpInt4 As Integer
Dim tmpLong1 As Long
Dim tmpLong2 As Long
Dim tmpLong3 As Long
Dim tmpLong4 As Long
Dim i As Long
Dim j As Long
Dim Break As Boolean
Dim lSnipCounter As Long
Dim lSafeCounter As Long
Dim bLanguage As Byte
Dim ScriptStructs() As tLevelScript
Dim lLevelCounter As Byte
Const lMaxLength As Long = 34
Dim iFileNum As Integer
Dim cString As cStringBuilder
Dim cData As cStringBuilder
Dim StdHeaderExists As Boolean
Dim bTemp() As Byte
Dim bTemp2() As Byte
Dim sTemp As String
    
    If lOffset = 0 Then
        Exit Function
    End If
    
    StartTiming
    
    On Error GoTo ErrHandler
    Set cString = New cStringBuilder
    Set cData = New cStringBuilder
    
    If Remove = 0 Then
        Screen.MousePointer = vbHourglass
        IsLoading = True
        LockUpdate Document(frmMain.Tabs.SelectedTab).hWnd
    End If
    
    EraseCol Snips
    EraseCol Strings
    EraseCol Moves
    EraseCol Marts
    EraseCol Brailles
    
    Snips.Add lOffset, CStr(lOffset)
      
    lSnipCounter = 1
    lSafeCounter = 0
    lLevelCounter = 0
    
    ReDim ScriptStructs(0) As tLevelScript
    
    If LenB(sCommentChar) <> 0 Then
        sCommentChar = LTrim$(sCommentChar)
    Else
        sCommentChar = "'"
    End If
    
    StdHeaderExists = FileExists(App.Path & "\std.rbh")
    
    iFileNum = FreeFile
    Open sFileName For Binary As #iFileNum
    
        Get #iFileNum, &HAC + 1, sGameCode
        Get #iFileNum, &HAF + 1, bLanguage
  
        Select Case sGameCode
            Case "AXV", "AXP"
                IsFireRed = False
                MaxCommand = &HC5
            Case "BPR", "BPG"
                IsFireRed = True
                MaxCommand = &HD4
            Case Else
                IsFireRed = False
                MaxCommand = &HE2
        End Select
  
        If bLanguage <> AscW("J") Then
            Japanese = False
        Else
            Japanese = True
        End If
        
        Do While lSnipCounter < Snips.Count + 1
          
            Break = False
            lSafeCounter = 0
            lOffset = Snips.Item(lSnipCounter)
            
            If Remove = 1 Then
                If lSnipCounter > 1 Then
                    Break = True
                    Exit Do
                End If
            End If
            
            If lOffset <= 0 Then Exit Do
            
            If iComments = 1 Then
                cString.Append sCommentChar & String$(15, "-") & vbNewLine
            End If
            
            cString.Append "#org" & sHexPrefix & Hex$(lOffset) & vbNewLine
            Seek #iFileNum, lOffset + 1
          
            Do
        
                Get #iFileNum, , tmpByte1
Begin:
                Select Case tmpByte1
                    
                    Case &H6, &H7
                      
                        If IsLevelScript = False Then
                        
                            Get #iFileNum, , tmpByte2
                            Get #iFileNum, , tmpLong2
                            
                            cString.Append "if" & sHexPrefix & Hex$(tmpByte2) & vbSpace
                            
                            If tmpByte1 = &H6 Then
                                cString.Append RubiCommands(&H5).Keyword
                            Else
                                cString.Append RubiCommands(&H4).Keyword
                            End If
                            
                            cString.Append sHexPrefix & Hex$(tmpLong2)
                            AddSnip (tmpLong2)
                        
                        Else
                            
                            cString.Append "#raw" & sHexPrefix & Hex$(tmpByte1) & vbNewLine
                            Get #iFileNum, , tmpLong2
                            
                            cString.Append "#raw pointer" & sHexPrefix & Hex$(tmpLong2)
                            AddSnip (tmpLong2)
                            
                        End If
                    
                    Case &H2
                        
                        If IsLevelScript = False Then
                            Break = True
                            cString.Append RubiCommands(tmpByte1).Keyword & vbNewLine
                        Else
                            
                            On Error GoTo FakeLevelScript
                            
                            cString.Append "#raw" & sHexPrefix & Hex$(tmpByte1) & vbNewLine
                            Get #iFileNum, , tmpLong2
                            
                            tmpLong1 = Loc(iFileNum) + 1
                            cString.Append "#raw pointer" & sHexPrefix & Hex$(tmpLong2)
                            
                            Get #iFileNum, (tmpLong2 - &H8000000 + 1), tmpInt1
                            Get #iFileNum, (tmpLong2 - &H8000000 + 1 + 2), tmpInt2
                            Get #iFileNum, (tmpLong2 - &H8000000 + 1 + 4), tmpLong3
                            Get #iFileNum, (tmpLong2 - &H8000000 + 1 + 8), tmpInt3
                            
                            If lLevelCounter > UBound(ScriptStructs) Then
                                ReDim Preserve ScriptStructs(lLevelCounter + 9)
                            End If
                            
                            ScriptStructs(lLevelCounter).Variable = tmpInt1
                            ScriptStructs(lLevelCounter).Value = tmpInt2
                            ScriptStructs(lLevelCounter).Pointer = tmpLong3
                            ScriptStructs(lLevelCounter).Offset = tmpLong2
                            ScriptStructs(lLevelCounter).Variable2 = tmpInt3
                            
                            If IsPtr(tmpLong3) Then
                                lLevelCounter = lLevelCounter + 1
                                AddSnip (tmpLong3)
                            End If
                            
                            Seek #iFileNum, tmpLong1
                            
                        End If
                        
                    Case &HF
                        
                        Get #iFileNum, , tmpByte2
                        Get #iFileNum, , tmpLong2
                        
                        If tmpByte2 = &H0 Then
                          
                            Get #iFileNum, , tmpByte3
                          
                            If tmpByte3 = &H9 Then
                          
                            cString.Append "msgbox" & sHexPrefix & Hex$(tmpLong2)
                            
                            Get #iFileNum, , tmpByte4
                            
                            If iDecompileMode <> Strict Then
                                                         
                                If StdHeaderExists Then
                            
                                    Select Case tmpByte4
                                        Case &H4: cString.Append " MSG_KEEPOPEN"
                                        Case &H6: cString.Append " MSG_NORMAL"
                                        Case &H3: cString.Append " MSG_SIGN"
                                        Case &H5: cString.Append " MSG_YESNO"
                                        Case &H2: cString.Append " MSG_FACE"
                                        Case Else: cString.Append sHexPrefix & Hex$(tmpByte4)
                                    End Select
                                
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpByte4)
                                End If
                                
                            Else
                                cString.Append sHexPrefix & Hex$(tmpByte4)
                            End If
                          
                            If IsPtr(tmpLong2) Then
                                
                                If iComments = 1 Then
                                    cString.Append vbSpace & sCommentChar & """StrRef" & Hex$((tmpLong2 And &HFFFFFFF) - &H8000000) & """"
                                End If
                                
                                AddString (tmpLong2)
                                
                            End If
                            
                          Else
                            
                            'loadpointer
                            
                            cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2) & sHexPrefix & Hex$(tmpLong2)
                            
                            If IsPtr(tmpLong2) Then
                                
                                If iComments = 1 Then
                                    cString.Append vbSpace & sCommentChar & """StrRef" & Hex$((tmpLong2 And &HFFFFFFF) - &H8000000) & """"
                                End If
                                
                                AddString (tmpLong2)
                                
                            End If
                            
                            cString.Append vbNewLine
                            
                            tmpByte1 = tmpByte3
                            GoTo Begin
                            
                          End If
                          
                        Else
                          cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2) & sHexPrefix & Hex$(tmpLong2)
                        End If
                        
                    Case &H3
                        
                        If IsLevelScript = False Then
                            Break = True
                            cString.Append RubiCommands(tmpByte1).Keyword & vbNewLine
                        Else
                        
                            cString.Append "#raw" & sHexPrefix & Hex$(tmpByte1) & vbNewLine
                            Get #iFileNum, , tmpLong2
                            
                            cString.Append "#raw pointer" & sHexPrefix & Hex$(tmpLong2)
                            AddSnip (tmpLong2)
                            
                        End If
                        
                    Case &H5
                        
                        If IsLevelScript = False Then
                            
                            Break = True
                            
                            Get #iFileNum, , tmpLong2
                            cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpLong2) & vbNewLine
                            AddSnip (tmpLong2)
                            
                            'Get #iFileNum, , tmpByte2
                            
                            'If tmpByte2 = &H2 Then
                            '    cString.Append RubiCommands(tmpByte2).Keyword & vbNewLine
                            'End If
                            
                        Else
                            
                            cString.Append "#raw" & sHexPrefix & Hex$(tmpByte1) & vbNewLine
                            Get #iFileNum, , tmpLong2
                            
                            cString.Append "#raw pointer" & sHexPrefix & Hex$(tmpLong2)
                            AddSnip (tmpLong2)
                            
                        End If
                        
                    Case &H4F, &H50
                        
                        Get #iFileNum, , tmpInt2
                        Get #iFileNum, , tmpLong2
                        
                        cString.Append RubiCommands(tmpByte1).Keyword
                        
                        If StdHeaderExists Then
                            
                            If iDecompileMode <> Strict Then
                            
                                Select Case tmpInt2
                                    Case &HFF: cString.Append " MOVE_PLAYER"
                                    Case &H800F: cString.Append " LASTTALKED"
                                    Case &H7F: cString.Append " MOVE_CAMERA"
                                    Case Else: cString.Append sHexPrefix & Hex$(tmpInt2)
                                End Select
                            
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                        
                        Else
                            cString.Append sHexPrefix & Hex$(tmpInt2)
                        End If
                                
                        cString.Append sHexPrefix & Hex$(tmpLong2)
                        
                        If tmpByte1 = &H50 Then
                            Get #iFileNum, , tmpByte2
                            Get #iFileNum, , tmpByte3
                            cString.Append sHexPrefix & Hex$(tmpByte2) & sHexPrefix & Hex$(tmpByte3)
                        End If
                        
                        AddMove (tmpLong2)
                        
                    Case &H4
                        
                        If IsLevelScript = False Then
                            Get #iFileNum, , tmpLong2
                            cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpLong2)
                            AddSnip (tmpLong2)
                        Else
                            
                            On Error GoTo FakeLevelScript
                            
                            cString.Append "#raw" & sHexPrefix & Hex$(tmpByte1) & vbNewLine
                            Get #iFileNum, , tmpLong2
                            
                            tmpLong1 = Loc(iFileNum) + 1
                            cString.Append "#raw pointer" & sHexPrefix & Hex$(tmpLong2)
                            
                            Get #iFileNum, (tmpLong2 - &H8000000 + 1), tmpInt1
                            Get #iFileNum, (tmpLong2 - &H8000000 + 1 + 2), tmpInt2
                            Get #iFileNum, (tmpLong2 - &H8000000 + 1 + 4), tmpLong3
                            Get #iFileNum, (tmpLong2 - &H8000000 + 1 + 8), tmpInt3
                            
                            ScriptStructs(lLevelCounter).Variable = tmpInt1
                            ScriptStructs(lLevelCounter).Value = tmpInt2
                            ScriptStructs(lLevelCounter).Pointer = tmpLong3
                            ScriptStructs(lLevelCounter).Offset = tmpLong2
                            ScriptStructs(lLevelCounter).Variable2 = tmpInt3
                            
                            If IsPtr(tmpLong3) Then
                                lLevelCounter = lLevelCounter + 1
                                AddSnip (tmpLong3)
                            End If
                            
                            Seek #iFileNum, tmpLong1
                            
                        End If
                        
                    Case &H5C 'trainerbattle
                        
                        Get #iFileNum, , tmpByte2
                        Get #iFileNum, , tmpInt2
                        Get #iFileNum, , tmpInt3
                        Get #iFileNum, , tmpLong2
                        
                        cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2) & sHexPrefix & Hex$(tmpInt2) & sHexPrefix & Hex$(tmpInt3) & sHexPrefix & Hex$(tmpLong2)
                        
                        If tmpByte2 <> &H3 Then
                        
                            Get #iFileNum, , tmpLong3
                            cString.Append sHexPrefix & Hex$(tmpLong3)
                            AddString (tmpLong2)
                            AddString (tmpLong3)
                            
                            Select Case tmpByte2
                                Case &H1, &H2
                                    Get #iFileNum, , tmpLong4
                                    cString.Append sHexPrefix & Hex$(tmpLong4)
                                    AddSnip (tmpLong4)
                                Case &H4, &H7
                                    Get #iFileNum, , tmpLong4
                                    cString.Append sHexPrefix & Hex$(tmpLong4)
                                    AddString (tmpLong4)
                                Case &H6, &H8
                                    Get #iFileNum, , tmpLong4
                                    cString.Append sHexPrefix & Hex$(tmpLong4)
                                    AddString (tmpLong4)
                                    Get #iFileNum, , tmpLong4
                                    cString.Append sHexPrefix & Hex$(tmpLong4)
                                    AddSnip (tmpLong4)
                            End Select
                        
                        Else
                            AddString (tmpLong2)
                        End If
                        
                    Case Is > MaxCommand
                        
                        Exit Do
                        
                    Case &H67, &H9B, &HBD, &HDB, &HDF 'msgboxes
                        
                        Get #iFileNum, , tmpLong2
                        
                        cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpLong2)
                        
                        If IsPtr(tmpLong2) Then
                            
                            If iComments = 1 Then
                                cString.Append vbSpace & sCommentChar & """StrRef" & Hex$((tmpLong2 And &HFFFFFFF) - &H8000000) & """"
                            End If
                            
                            AddString (tmpLong2)
                            
                        End If
                    
                    Case &H0
                        
                        If IsLevelScript = False Then
                            cString.Append RubiCommands(tmpByte1).Keyword
                            lSafeCounter = lSafeCounter + 1
                        Else
                            
                            Break = True
                            IsLevelScript = False
                    
                            cString.Append "#raw" & sHexPrefix & Hex$(tmpByte1) & vbNewLine
                      
                            For i = 0 To lLevelCounter - 1
                            
                                cString.Append vbNewLine
                                
                                If iComments = 1 Then
                                    cString.Append sCommentChar & "---------------" & vbNewLine
                                End If
                                
                                cString.Append "#org" & sHexPrefix & Hex$((ScriptStructs(i).Offset And &HFFFFFFF) - &H8000000) & vbNewLine
                                cString.Append "#raw word" & sHexPrefix & Hex$(ScriptStructs(i).Variable) & vbNewLine
                                cString.Append "#raw word" & sHexPrefix & Hex$(ScriptStructs(i).Value) & vbNewLine
                                cString.Append "#raw pointer" & sHexPrefix & Hex$(ScriptStructs(i).Pointer) & vbNewLine
                                cString.Append "#raw word" & sHexPrefix & Hex$(ScriptStructs(i).Variable2) & vbNewLine
                                
                            Next i
                      
                        End If
                        
                    Case &H1A
                        
                        Get #iFileNum, , tmpInt2
                
                        If tmpInt2 = &H8000 Then
                          
                            Get #iFileNum, , tmpInt2
                            Get #iFileNum, , tmpByte1
                          
                            If tmpByte1 <> &H9 Then
                            
                                Get #iFileNum, , tmpInt3
                                Get #iFileNum, , tmpInt4
                                Get #iFileNum, , tmpByte2
                                Get #iFileNum, , tmpByte3
                            
                                If tmpByte2 = &H9 Then
                                    
                                    cString.Append "giveitem"
                                    
                                    If tmpInt2 <> &H800D Then
                                        If iDecompileMode <> Enhanced Then
                                            cString.Append sHexPrefix & Hex$(tmpInt2)
                                        Else
                                            If tmpInt2 >= 0 And tmpInt2 <= &H178 Then
                                                cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt2)
                                                If cString.Find("ITEM_") <> 0 Then
                                                    If cString.Find("#include stditems.rbh") = 0 Then
                                                        cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                                    End If
                                                End If
                                            Else
                                                cString.Append sHexPrefix & Hex$(tmpInt2)
                                            End If
                                        End If
                                    Else
                                        If StdHeaderExists Then
                                            If iDecompileMode <> Strict Then
                                                cString.Append " LASTRESULT"
                                            Else
                                                cString.Append sHexPrefix & Hex$(tmpInt2)
                                            End If
                                        Else
                                            cString.Append sHexPrefix & Hex$(tmpInt2)
                                        End If
                                    End If
                                    
                                    cString.Append sHexPrefix & Hex$(tmpInt4)
                                    
                                    If StdHeaderExists Then
                                    
                                        If iDecompileMode <> Strict Then
                                    
                                            Select Case tmpByte3
                                                Case &H1: cString.Append " MSG_FIND"
                                                Case &H0: cString.Append " MSG_OBTAIN"
                                                Case Else: cString.Append sHexPrefix & Hex$(tmpByte3)
                                            End Select
                                        
                                        Else
                                            cString.Append sHexPrefix & Hex$(tmpByte3)
                                        End If
                                    
                                    Else
                                        cString.Append sHexPrefix & Hex$(tmpByte3)
                                    End If
                                    
                                ElseIf tmpByte2 = &H1A Then
                                    Get #iFileNum, , tmpByte3
                                    Get #iFileNum, , tmpInt3
                                    Get #iFileNum, , tmpByte1
                                    Get #iFileNum, , tmpByte2
                                    
                                    cString.Append "giveitem2"
                                    
                                    If tmpInt2 <> &H800D Then
                                        If iDecompileMode <> Enhanced Then
                                            cString.Append sHexPrefix & Hex$(tmpInt2)
                                        Else
                                            If tmpInt2 >= 0 And tmpInt2 <= &H178 Then
                                                cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt2)
                                                If cString.Find("ITEM_") <> 0 Then
                                                    If cString.Find("#include stditems.rbh") = 0 Then
                                                        cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                                    End If
                                                End If
                                            Else
                                                cString.Append sHexPrefix & Hex$(tmpInt2)
                                            End If
                                        End If
                                    Else
                                        If StdHeaderExists Then
                                            If iDecompileMode <> Strict Then
                                                cString.Append " LASTRESULT"
                                            Else
                                                cString.Append sHexPrefix & Hex$(tmpInt2)
                                            End If
                                        Else
                                            cString.Append sHexPrefix & Hex$(tmpInt2)
                                        End If
                                    End If
                                    
                                    cString.Append sHexPrefix & Hex$(tmpInt4) & sHexPrefix & Hex$(tmpInt3)
                                    
                                End If
                                
                            Else
                                
                                Get #iFileNum, , tmpByte2
                                
                                If tmpByte2 = &H7 Then
                                
                                    cString.Append "giveitem3"
                                    
                                    If iDecompileMode <> Enhanced Then
                                        cString.Append sHexPrefix & Hex$(tmpInt2)
                                    Else
                                        If tmpInt2 >= 0 And tmpInt2 <= &H178 Then
                                            cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt2 + &H179)
                                            If cString.Find("ITEM_") <> 0 Then
                                                If cString.Find("#include stditems.rbh") = 0 Then
                                                    cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                                End If
                                            End If
                                        Else
                                            cString.Append sHexPrefix & Hex$(tmpInt2)
                                        End If
                                    End If
                                    
                                ElseIf tmpByte2 = &H8 Then
                                    cString.Append "registernav" & sHexPrefix & Hex$(tmpInt2)
                                End If
                                
                            End If
                          
                        Else
                            Get #iFileNum, , tmpInt3
                            cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpInt2) & sHexPrefix & Hex$(tmpInt3)
                        End If
                        
                    Case &H7C 'checkattack
                        
                        Get #iFileNum, , tmpInt2
                        
                        cString.Append RubiCommands(tmpByte1).Keyword
                        
                        If iDecompileMode <> Enhanced Then
                            cString.Append sHexPrefix & Hex$(tmpInt2)
                        Else
                            If tmpInt2 >= 0 And tmpInt2 <= &H162 Then
                                cString.Append GetHeaderDefine(App.Path & "\stdattacks.rbh", tmpInt2)
                                If cString.Find("ATK_") <> 0 Then
                                    If cString.Find("#include stdattacks.rbh") = 0 Then
                                        cString.Insert 0, "#include stdattacks.rbh" & vbNewLine
                                    End If
                                End If
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                        End If
                        
                    Case &H82 'bufferattack
                        
                        Get #iFileNum, , tmpByte2
                        Get #iFileNum, , tmpInt2
                        
                        cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2)
                        
                        If iDecompileMode <> Enhanced Then
                            cString.Append sHexPrefix & Hex$(tmpInt2)
                        Else
                            If tmpInt2 >= 0 And tmpInt2 <= &H162 Then
                                cString.Append GetHeaderDefine(App.Path & "\stdattacks.rbh", tmpInt2)
                                If cString.Find("ATK_") <> 0 Then
                                    If cString.Find("#include stdattacks.rbh") = 0 Then
                                        cString.Insert 0, "#include stdattacks.rbh" & vbNewLine
                                    End If
                                End If
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                        End If
                        
                    Case &H44, &H45, &H46, &H47, &H49, &H4A
                        
                        Get #iFileNum, , tmpInt2
                        Get #iFileNum, , tmpInt3
                        
                        cString.Append RubiCommands(tmpByte1).Keyword
                        
                        If iDecompileMode <> Enhanced Then
                            cString.Append sHexPrefix & Hex$(tmpInt2)
                        Else
                            If tmpInt2 >= 0 And tmpInt2 <= &H178 Then
                                cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt2)
                                If cString.Find("ITEM_") <> 0 Then
                                    If cString.Find("#include stditems.rbh") = 0 Then
                                        cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                    End If
                                End If
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                        End If
                        
                        cString.Append sHexPrefix & Hex$(tmpInt3)
                        
                    Case &H78 'braille
                        
                        Get #iFileNum, , tmpLong2
                        
                        cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpLong2)
                        
                        If IsPtr(tmpLong2) Then
                            
                            If iComments = 1 Then
                                cString.Append vbSpace & sCommentChar & """BraRef" & Hex$((tmpLong2 And &HFFFFFFF) - &H8000000) & """"
                            End If
                            
                            AddBraille (tmpLong2)
                            
                        End If
                        
                    Case &H7D 'bufferpokemon
                        
                        Get #iFileNum, , tmpByte2
                        Get #iFileNum, , tmpInt2
                        
                        cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2)
                        
                        If tmpInt2 <> &H800D Then
                        
                            If iDecompileMode <> Enhanced Then
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            Else
                                If tmpInt2 >= 0 And tmpInt2 <= 411 Then
                                    cString.Append GetHeaderDefine(App.Path & "\stdpoke.rbh", tmpInt2)
                                    If cString.Find("PKMN_") <> 0 Then
                                        If cString.Find("#include stdpoke.rbh") = 0 Then
                                            cString.Insert 0, "#include stdpoke.rbh" & vbNewLine
                                        End If
                                    End If
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt2)
                                End If
                            End If
                            
                        Else
                            
                            If StdHeaderExists Then
                                If iDecompileMode <> Strict Then
                                    cString.Append " LASTRESULT"
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt2)
                                End If
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                            
                        End If
                        
                    Case &HA1 'cry
                        
                        Get #iFileNum, , tmpInt2
                        Get #iFileNum, , tmpInt3
                        
                        cString.Append RubiCommands(tmpByte1).Keyword
                        
                        If tmpInt2 <> &H800D Then
                            
                            If iDecompileMode <> Enhanced Then
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            Else
                                If tmpInt2 >= 0 And tmpInt2 <= 411 Then
                                    cString.Append GetHeaderDefine(App.Path & "\stdpoke.rbh", tmpInt2)
                                    If cString.Find("PKMN_") <> 0 Then
                                        If cString.Find("#include stdpoke.rbh") = 0 Then
                                            cString.Insert 0, "#include stdpoke.rbh" & vbNewLine
                                        End If
                                    End If
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt2)
                                End If
                            End If
                            
                        Else
                            
                            If StdHeaderExists Then
                                If iDecompileMode <> Strict Then
                                    cString.Append " LASTRESULT"
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt2)
                                End If
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                            
                        End If
                        
                        cString.Append sHexPrefix & Hex$(tmpInt3)
                      
                    Case &H1
                        
                        If IsLevelScript = False Then
                            cString.Append RubiCommands(tmpByte1).Keyword
                        Else
                            
                            cString.Append "#raw" & sHexPrefix & Hex$(tmpByte1) & vbNewLine
                            Get #iFileNum, , tmpLong2
                            
                            cString.Append "#raw pointer" & sHexPrefix & Hex$(tmpLong2)
                            AddSnip (tmpLong2)
                            
                        End If
                        
                    Case &H80 'bufferitem
                        
                        Get #iFileNum, , tmpByte2
                        Get #iFileNum, , tmpInt2
                        
                        cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2)
                        
                        If iDecompileMode <> Enhanced Then
                            cString.Append sHexPrefix & Hex$(tmpInt2)
                        Else
                            If tmpInt2 >= 0 And tmpInt2 <= &H178 Then
                                cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt2)
                                If cString.Find("ITEM_") <> 0 Then
                                    If cString.Find("#include stditems.rbh") = 0 Then
                                        cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                    End If
                                End If
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                        End If
                        
                    Case &H86, &H87, &H88 'pokemart
                        
                        Get #iFileNum, , tmpLong2
                        cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpLong2)
                        
                        AddMart (tmpLong2)
                        
                    Case &H75
                        
                        Get #iFileNum, , tmpInt2
                        Get #iFileNum, , tmpByte2
                        Get #iFileNum, , tmpByte3
                        
                        cString.Append RubiCommands(tmpByte1).Keyword
                        
                        If iDecompileMode <> Enhanced Then
                            cString.Append sHexPrefix & Hex$(tmpInt2)
                        Else
                            If tmpInt2 >= 0 And tmpInt2 <= 411 Then
                                cString.Append GetHeaderDefine(App.Path & "\stdpoke.rbh", tmpInt2)
                                If cString.Find("PKMN_") <> 0 Then
                                    If cString.Find("#include stdpoke.rbh") = 0 Then
                                        cString.Insert 0, "#include stdpoke.rbh" & vbNewLine
                                    End If
                                End If
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                        End If
                        
                        cString.Append sHexPrefix & Hex$(tmpByte2) & sHexPrefix & Hex$(tmpByte3)
                        
                    
                    Case &H79, &H7A 'givepokemon, giveegg
                        
                        Get #iFileNum, , tmpInt2

                        cString.Append RubiCommands(tmpByte1).Keyword
                        
                        If iDecompileMode <> Enhanced Then
                            cString.Append sHexPrefix & Hex$(tmpInt2)
                        Else
                            If tmpInt2 >= 0 And tmpInt2 <= 411 Then
                                cString.Append GetHeaderDefine(App.Path & "\stdpoke.rbh", tmpInt2)
                                If cString.Find("PKMN_") <> 0 Then
                                    If cString.Find("#include stdpoke.rbh") = 0 Then
                                        cString.Insert 0, "#include stdpoke.rbh" & vbNewLine
                                    End If
                                End If
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                        End If
                        
                        If tmpByte1 = &H79 Then
                            
                            Get #iFileNum, , tmpByte2
                            Get #iFileNum, , tmpInt3
                            Get #iFileNum, , tmpLong1
                            Get #iFileNum, , tmpLong2
                            Get #iFileNum, , tmpByte3
                            
                            cString.Append sHexPrefix & Hex$(tmpByte2)
                            
                            If iDecompileMode <> Enhanced Then
                                cString.Append sHexPrefix & Hex$(tmpInt3)
                            Else
                                If tmpInt3 >= 0 And tmpInt3 <= &H178 Then
                                    cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt3)
                                    If cString.Find("ITEM_") <> 0 Then
                                        If cString.Find("#include stditems.rbh") = 0 Then
                                            cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                        End If
                                    End If
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt3)
                                End If
                            End If
                            
                            cString.Append sHexPrefix & Hex$(tmpLong1) & sHexPrefix & Hex$(tmpLong2) & sHexPrefix & Hex$(tmpByte3)
                        End If
                        
                    Case &HB6 'wildbattle
                        
                        Get #iFileNum, , tmpInt2
                        Get #iFileNum, , tmpByte2
                        Get #iFileNum, , tmpInt3
                        Get #iFileNum, , tmpByte3
                        
                        If tmpByte3 = &HB7 Then
                        
                            cString.Append "wildbattle"
                            
                            If iDecompileMode <> Enhanced Then
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            Else
                                If tmpInt2 >= 0 And tmpInt2 <= 411 Then
                                    cString.Append GetHeaderDefine(App.Path & "\stdpoke.rbh", tmpInt2)
                                    If cString.Find("PKMN_") <> 0 Then
                                        If cString.Find("#include stdpoke.rbh") = 0 Then
                                            cString.Insert 0, "#include stdpoke.rbh" & vbNewLine
                                        End If
                                    End If
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt2)
                                End If
                            End If
                            
                            cString.Append sHexPrefix & Hex$(tmpByte2)
                            
                            If iDecompileMode <> Enhanced Then
                                cString.Append sHexPrefix & Hex$(tmpInt3)
                            Else
                                If tmpInt3 >= 0 And tmpInt3 <= &H178 Then
                                    cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt3)
                                    If cString.Find("ITEM_") <> 0 Then
                                        If cString.Find("#include stditems.rbh") = 0 Then
                                            cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                        End If
                                    End If
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt3)
                                End If
                            End If
                            
                        ElseIf tmpByte3 = &H25 Then
                            
                            Get #iFileNum, , tmpInt4
                            
                            cString.Append "wildbattle2"
                            
                            If iDecompileMode <> Enhanced Then
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            Else
                                If tmpInt2 >= 0 And tmpInt2 <= 411 Then
                                    cString.Append GetHeaderDefine(App.Path & "\stdpoke.rbh", tmpInt2)
                                    If cString.Find("PKMN_") <> 0 Then
                                        If cString.Find("#include stdpoke.rbh") = 0 Then
                                            cString.Insert 0, "#include stdpoke.rbh" & vbNewLine
                                        End If
                                    End If
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt2)
                                End If
                            End If
                            
                            cString.Append sHexPrefix & Hex$(tmpByte2)
                            
                            If iDecompileMode <> Enhanced Then
                                cString.Append sHexPrefix & Hex$(tmpInt3)
                            Else
                                If tmpInt3 >= 0 And tmpInt3 <= &H178 Then
                                    cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt3)
                                    If cString.Find("ITEM_") <> 0 Then
                                        If cString.Find("#include stditems.rbh") = 0 Then
                                            cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                        End If
                                    End If
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt3)
                                End If
                            End If
                            
                            Select Case tmpInt4
                                Case &H137: cString.Append sHexPrefix & "0"
                                Case &H138: cString.Append sHexPrefix & "1"
                                Case &H139: cString.Append sHexPrefix & "2"
                                Case &H13A: cString.Append sHexPrefix & "3"
                                Case &H13B: cString.Append sHexPrefix & "4"
                                Case &H143: cString.Append sHexPrefix & "5"
                                Case &H156: cString.Append sHexPrefix & "6"
                                Case Else: cString.Append sHexPrefix & "0"
                            End Select
                            
                            Get #iFileNum, , tmpByte4
                            
                        Else
                            
                            cString.Append RubiCommands(tmpByte1).Keyword
                            
                            If iDecompileMode <> Enhanced Then
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            Else
                                If tmpInt2 >= 0 And tmpInt2 <= 411 Then
                                    cString.Append GetHeaderDefine(App.Path & "\stdpoke.rbh", tmpInt2)
                                    If cString.Find("PKMN_") <> 0 Then
                                        If cString.Find("#include stdpoke.rbh") = 0 Then
                                            cString.Insert 0, "#include stdpoke.rbh" & vbNewLine
                                        End If
                                    End If
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt2)
                                End If
                            End If
                            
                            cString.Append sHexPrefix & Hex$(tmpByte2)
                            
                            If iDecompileMode <> Enhanced Then
                                cString.Append sHexPrefix & Hex$(tmpInt3)
                            Else
                                If tmpInt3 >= 0 And tmpInt3 <= &H178 Then
                                    cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt3)
                                    If cString.Find("ITEM_") <> 0 Then
                                        If cString.Find("#include stditems.rbh") = 0 Then
                                            cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                        End If
                                    End If
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt3)
                                End If
                            End If
                            
                            cString.Append vbNewLine
                            
                            tmpByte1 = tmpByte3
                            GoTo Begin
                            
                        End If
                     
                    Case &H48, &H4B, &H4C, &H4D, &H4E
                        
                        Get #iFileNum, , tmpInt2
                        
                        cString.Append RubiCommands(tmpByte1).Keyword
                        
                        If iDecompileMode <> Enhanced Then
                            cString.Append sHexPrefix & Hex$(tmpInt2)
                        Else
                            If tmpInt2 >= 0 And tmpInt2 <= &H178 Then
                                cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt2)
                                If cString.Find("ITEM_") <> 0 Then
                                    If cString.Find("#include stditems.rbh") = 0 Then
                                        cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                    End If
                                End If
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                        End If

                    Case &H5E
                      
                      Break = True
                      cString.Append RubiCommands(tmpByte1).Keyword & vbNewLine

                    Case &H81 'bufferdecoration
                        
                        Get #iFileNum, , tmpByte2
                        Get #iFileNum, , tmpInt2
                        
                        cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2)
                        
                        If iDecompileMode <> Enhanced Then
                            cString.Append sHexPrefix & Hex$(tmpInt2)
                        Else
                            If tmpInt2 > 0 And tmpInt2 <= &H78 Then
                                cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt2 + &H179)
                                If cString.Find("ITEM_") <> 0 Then
                                    If cString.Find("#include stditems.rbh") = 0 Then
                                        cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                    End If
                                End If
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                        End If

                    Case &H85, &HBF
                        
                        Get #iFileNum, , tmpByte2
                        Get #iFileNum, , tmpLong2
                        
                        cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2) & sHexPrefix & Hex$(tmpLong2)
                        
                        If IsPtr(tmpLong2) Then
                            
                            If iComments = 1 Then
                                cString.Append vbSpace & sCommentChar & """StrRef" & Hex$((tmpLong2 And &HFFFFFFF) - &H8000000) & """"
                            End If
                            
                            AddString (tmpLong2)
                            
                        End If
                      
                    Case &HB9, &HBA
                        
                        Get #iFileNum, , tmpLong2
                        cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpLong2)
                        AddSnip (tmpLong2)
                      
                    Case &HBB, &HBC
                      
                      Get #iFileNum, , tmpByte2
                      Get #iFileNum, , tmpLong2
                      
                      cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2) & sHexPrefix & Hex$(tmpLong2)
                      AddSnip (tmpLong2)
                    
                    Case &H8
                    
                      Break = True
                      Get #iFileNum, , tmpByte2
                      cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2) & vbNewLine
                      
                    Case &HD
                      
                      Break = True
                      cString.Append RubiCommands(tmpByte1).Keyword & vbNewLine
                      
                    Case &H24
                        
                      Break = True
                      Get #iFileNum, , tmpLong2
                      cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpLong2) & vbNewLine
                    
                    Case &HD3
                        
                        If IsFireRed = False Then
                            Get #iFileNum, , tmpInt2
                            cString.Append "cmd" & Hex$(tmpByte1) & sHexPrefix & Hex$(tmpInt2)
                        Else
                            GoTo Continue
                        End If
                      
                    Case &HD4
                        
                        If IsFireRed = False Then
                            cString.Append "cmd" & Hex$(tmpByte1)
                        Else
                            
                            Get #iFileNum, , tmpByte2
                            Get #iFileNum, , tmpInt2
                            Get #iFileNum, , tmpInt3
                        
                            cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2)
                            
                            If iDecompileMode <> Enhanced Then
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            Else
                                If tmpInt2 >= 0 And tmpInt2 <= &H178 Then
                                    cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt2)
                                    If cString.Find("ITEM_") <> 0 Then
                                        If cString.Find("#include stditems.rbh") = 0 Then
                                            cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                        End If
                                    End If
                                Else
                                    cString.Append sHexPrefix & Hex$(tmpInt2)
                                End If
                            End If
                            
                            cString.Append sHexPrefix & Hex$(tmpInt3)
                        
                        End If
                        
                    Case &HE2
                    
                        Get #iFileNum, , tmpByte2
                        Get #iFileNum, , tmpInt2
                        Get #iFileNum, , tmpInt3
                    
                        cString.Append RubiCommands(tmpByte1).Keyword & sHexPrefix & Hex$(tmpByte2)
                        
                        If iDecompileMode <> Enhanced Then
                            cString.Append sHexPrefix & Hex$(tmpInt2)
                        Else
                            If tmpInt2 >= 0 And tmpInt2 <= &H178 Then
                                cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", tmpInt2)
                                If cString.Find("ITEM_") <> 0 Then
                                    If cString.Find("#include stditems.rbh") = 0 Then
                                        cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                    End If
                                End If
                            Else
                                cString.Append sHexPrefix & Hex$(tmpInt2)
                            End If
                        End If
                        
                        cString.Append sHexPrefix & Hex$(tmpInt3)
                
                    Case Else
Continue:
                        cString.Append RubiCommands(tmpByte1).Keyword
                        
                        If RubiCommands(tmpByte1).ParamCount > 0 Then
                            
                            For j = 0 To RubiCommands(tmpByte1).ParamCount - 1
                                
                                Select Case RubiParams(tmpByte1, j).SIZE
                                    
                                    Case 1 'Word
                                    
                                        Get #iFileNum, , tmpInt2
                                        
                                        If StdHeaderExists Then
                                            
                                            If iDecompileMode <> Strict Then
                                            
                                                Select Case tmpInt2
                                                    
                                                    Case Is > &H800F '-32753
                                                        cString.Append sHexPrefix & Hex$(tmpInt2)
                                                    
                                                    Case &H8000 To &H800B, &H800E
                                                        cString.Append sHexPrefix & Hex$(tmpInt2)
                                                        
                                                    Case &H800D, &H800F, &H800C
                                                    
                                                        If tmpInt2 = &H800D Then
                                                            cString.Append " LASTRESULT"
                                                        ElseIf tmpInt2 = &H800F Then
                                                            cString.Append " LASTTALKED"
                                                        Else
                                                            cString.Append " PLAYERFACING"
                                                        End If
                                                        
                                                    Case Else
                                                        cString.Append sHexPrefix & Hex$(tmpInt2)
                                                        
                                                End Select
                                            
                                            Else
                                                cString.Append sHexPrefix & Hex$(tmpInt2)
                                            End If
                                        
                                        Else
                                            cString.Append sHexPrefix & Hex$(tmpInt2)
                                        End If
                                        
                                        'cString.Append sHexPrefix & Hex$(tmpInt2)
                                        
                                    Case 0 'Byte
                                    
                                        Get #iFileNum, , tmpByte2
                                        cString.Append sHexPrefix & Hex$(tmpByte2)
                                        
                                    Case 3, 2 'Pointer, DWord
                                        
                                        Get #iFileNum, , tmpLong2
                                        cString.Append sHexPrefix & Hex$(tmpLong2)
                                      
                              End Select
                              
                            Next j
                            
                        End If
                      
                End Select
                
                If lSafeCounter > 14 Then Exit Do
                
                If Loc(iFileNum) >= LOF(iFileNum) Then
                    Exit Do
                End If
                
                cString.Append vbNewLine
                
            Loop Until Break = True
            
            If Remove = 2 Then
                
                If Break = True Then
                    
                    If Loc(iFileNum) > CLng(Snips.Item(lSnipCounter)) Then
                    
                        ReDim bTemp((Loc(iFileNum) - CLng(Snips.Item(lSnipCounter)) - 1&)) As Byte
                                            
                        If bFreeSpace <> 0 Then
                            RtlFillMemory bTemp(0), UBound(bTemp) + 1&, bFreeSpace
                        End If
                        
                        Put #iFileNum, CLng(Snips.Item(lSnipCounter)) + 1&, bTemp
                        Erase bTemp
                    
                    End If
                
                End If
                
            End If
            
            lSnipCounter = lSnipCounter + 1
            
        Loop
  
        Decompile = Loc(iFileNum)
        
        If Remove = 1 Then
            If Break = False Then
                Decompile = 0
            End If
            Exit Function
        End If
  
        If Strings.Count > 0 Then
          
            If iComments = 1 Then
                cString.Append vbNewLine & sCommentChar & "---------"
                cString.Append vbNewLine & sCommentChar & " Strings"
                cString.Append vbNewLine & sCommentChar & "---------" & vbNewLine
            End If
          
            For i = 1 To Strings.Count
              
                cString.Append "#org" & sHexPrefix & Hex$(Strings.Item(i)) & vbNewLine
                tmpLong1 = Strings.Item(i)
                
                If tmpLong1 > 0 Then
                    
                    ReDim bTemp(999) As Byte
                    Get #iFileNum, tmpLong1 + 1, bTemp
                    
                    If Remove = 0 Then
                    
                        sTemp = Sapp2Asc(bTemp, Japanese)
                        cString.Append "= " & Sapp2Asc(bTemp, Japanese) & vbNewLine & vbNewLine
                        
                        Erase bTemp
                    
                        If iComments = 1 Then
                            
                            Do
                            
                                tmpLong2 = cString.Find("StrRef" & Hex$(Strings.Item(i)))
                                
                                If tmpLong2 <> 0 Then
                                
                                    tmpLong2 = tmpLong2 - 1
                                
                                    cString.Remove tmpLong2, Len("StrRef" & Hex$(Strings.Item(i)))
                                    
                                    If Len(sTemp) > lMaxLength Then
                                        cString.Insert tmpLong2, LeftB$(sTemp, lMaxLength * 2) & "..."
                                    Else
                                        cString.Insert tmpLong2, sTemp
                                    End If
                                    
                                End If
                            
                            Loop Until tmpLong2 = 0
                            
                        End If
                        
                    ElseIf Remove = 2 Then
                        
                        ReDim bTemp2(0) As Byte: bTemp2(0) = bFreeSpace
                        tmpLong2 = InStrB(bTemp, bTemp2)
                        
                        If tmpLong2 <> 0 Then
                            If bTemp(tmpLong2) = 0 Then
                                tmpLong2 = tmpLong2 + 1
                            End If
                        End If
                        
                        If tmpLong2 > 1 Then
                        
                            ReDim bTemp2(tmpLong2 - 1) As Byte
                                                
                            If bFreeSpace <> 0 Then
                                RtlFillMemory bTemp2(0), UBound(bTemp2) + 1&, bFreeSpace
                            End If
                            
                            Put #iFileNum, tmpLong1 + 1&, bTemp2
                            Erase bTemp2
                            
                        End If
                        
                    End If
                    
                End If
              
            Next i
        
        End If
    
        If Brailles.Count > 0 Then
            
            If iComments = 1 Then
                cString.Append vbNewLine & sCommentChar & "---------"
                cString.Append vbNewLine & sCommentChar & " Braille"
                cString.Append vbNewLine & sCommentChar & "---------" & vbNewLine
            End If
      
            For i = 1 To Brailles.Count
            
                cString.Append "#org" & sHexPrefix & Hex$(Brailles.Item(i)) & vbNewLine
                tmpLong1 = Brailles.Item(i)
                
                If tmpLong1 > 0 Then
                    
                    ReDim bTemp(399) As Byte
                    Get #iFileNum, tmpLong1 + 1, bTemp
                    
                    If cString.Find(RubiCommands(&H73).Keyword) > 0 Then
              
                        cString.Append "#raw"
              
                        For j = 0 To 5
                            cString.Append sHexPrefix & Hex$(bTemp(j))
                        Next j
                        
                        For j = LBound(bTemp) To UBound(bTemp) - 6
                            bTemp(j) = bTemp(j + 6)
                        Next j
                        
                        ReDim Preserve bTemp(UBound(bTemp) - 6) As Byte
                        cString.Append vbNewLine
                        
                    End If
                    
                    If Remove = 0 Then
                    
                        sTemp = Braille2Asc(bTemp)
                        cString.Append "#braille " & sTemp & vbNewLine & vbNewLine
                        Erase bTemp
                    
                        If iComments = 1 Then
                            
                            Do
                            
                                tmpLong2 = cString.Find("BraRef" & Hex$(Brailles.Item(i)))
                                
                                If tmpLong2 <> 0 Then
                                
                                    tmpLong2 = tmpLong2 - 1
                                
                                    cString.Remove tmpLong2, Len("BraRef" & Hex$(Brailles.Item(i)))
                                    
                                    If Len(sTemp) <= lMaxLength Then
                                        cString.Insert tmpLong2, sTemp
                                    Else
                                        cString.Insert tmpLong2, LeftB$(sTemp, lMaxLength * 2) & "..."
                                    End If
                                    
                                End If
                            
                            Loop Until tmpLong2 = 0
                                
                        End If

                    ElseIf Remove = 2 Then
                        
                        ReDim bTemp2(0) As Byte: bTemp2(0) = bFreeSpace
                        tmpLong2 = InStrB(bTemp, bTemp2)
                        
                        If tmpLong2 <> 0 Then
                            If bTemp(tmpLong2) = 0 Then
                                tmpLong2 = tmpLong2 + 1
                            End If
                        End If
                        
                        If tmpLong2 > 1 Then
                        
                            ReDim bTemp2(tmpLong2 - 1) As Byte
                                                
                            If bFreeSpace <> 0 Then
                                RtlFillMemory bTemp2(0), UBound(bTemp2) + 1&, bFreeSpace
                            End If
                            
                            Put #iFileNum, tmpLong1 + 1&, bTemp2
                            Erase bTemp2
                            
                        End If
                        
                    End If
                    
                End If
            
            Next i
        
        End If
        
        If Moves.Count > 0 Then
            
            If iComments = 1 Then
                cString.Append vbNewLine & sCommentChar & "-----------"
                cString.Append vbNewLine & sCommentChar & " Movements"
                cString.Append vbNewLine & sCommentChar & "-----------" & vbNewLine
            End If
        
            For i = 1 To Moves.Count
          
                cString.Append "#org" & sHexPrefix & Hex$(Moves.Item(i)) & vbNewLine
                tmpLong1 = Moves.Item(i)
                
                If tmpLong1 > 0 Then
                    
                    ReDim bTemp(511) As Byte
                    Get #iFileNum, tmpLong1 + 1, bTemp
                    
                    For j = LBound(bTemp) To UBound(bTemp)
                        
                        If bTemp(j) <> &HFF Then
                            
                            If Remove = 0 Then
                                If iComments = 1 Then
                                    cString.Append "#raw" & sHexPrefix & Hex$(bTemp(j)) & vbSpace & sCommentChar & GetMovementLabel(bTemp(j)) & vbNewLine
                                Else
                                    cString.Append "#raw" & sHexPrefix & Hex$(bTemp(j)) & vbNewLine
                                End If
                            End If
                            
                            If bTemp(j) = &HFE Then
                                j = j + 1
                                Exit For
                            End If
                            
                        Else
                            Exit For
                        End If
                    
                    Next j
                    
                    If Remove = 0 Then
                        
                        Erase bTemp
                        cString.Append vbNewLine
                        
                    ElseIf Remove = 2 Then
                    
                        If j > 0 Then

                            ReDim bTemp2(j - 1) As Byte
                                                
                            If bFreeSpace <> 0 Then
                                RtlFillMemory bTemp2(0), UBound(bTemp2) + 1&, bFreeSpace
                            End If
                            
                            Put #iFileNum, tmpLong1 + 1&, bTemp2
                            Erase bTemp2
                        
                        End If
                        
                    End If
                        
                End If
                        
            Next i
        
        End If
    
        If Marts.Count > 0 Then
            
            If iComments = 1 Then
                cString.Append vbNewLine & sCommentChar & "-----------"
                cString.Append vbNewLine & sCommentChar & " MartItems"
                cString.Append vbNewLine & sCommentChar & "-----------" & vbNewLine
            End If
            
            Dim iTemp(99) As Integer
      
            For i = 1 To Marts.Count
            
                cString.Append "#org" & sHexPrefix & Hex$(Marts.Item(i)) & vbNewLine
                tmpLong1 = Marts.Item(i)
                
                If tmpLong1 > 0 Then
                    
                    Get #iFileNum, tmpLong1 + 1, iTemp
                    
                    For j = LBound(iTemp) To UBound(iTemp)
                    
                        If iTemp(j) <= &H178& Then
                            
                            If iTemp(j) >= 0 Then
                            
                                cString.Append "#raw word"
                            
                                If iDecompileMode <> Enhanced Then
                                    cString.Append sHexPrefix & Hex$(iTemp(j))
                                Else
                                    cString.Append GetHeaderDefine(App.Path & "\stditems.rbh", iTemp(j))
                                    If cString.Find("ITEM_") <> 0 Then
                                        If cString.Find("#include stditems.rbh") = 0 Then
                                            cString.Insert 0, "#include stditems.rbh" & vbNewLine
                                         End If
                                    End If
                                End If
                        
                                cString.Append vbNewLine
                                
                                If iTemp(j) = &H0 Then
                                    j = j + 1
                                    Exit For
                                End If
                                
                            Else
                                Exit For
                            End If
                            
                        Else
                            Exit For
                        End If
                    
                    Next j
                    
                    If Remove = 0 Then
                        
                        cString.Append vbNewLine
                    
                    ElseIf Remove = 2 Then
                    
                        If j > 0 Then

                            ReDim bTemp2(j * 2& - 1&) As Byte
                                                
                            If bFreeSpace <> 0 Then
                                RtlFillMemory bTemp2(0), UBound(bTemp2) + 1&, bFreeSpace
                            End If
                            
                            Put #iFileNum, tmpLong1 + 1&, bTemp2
                            Erase bTemp2
                        
                        End If
                        
                    End If
                    
                End If
                
            Next i
        
        End If
      
    Close #iFileNum
    
    If Remove = 2 Then
        Exit Function
    End If
    
    ' Remove the trailing vbNewLines
    For i = 1 To 2
        If Asc(RightB$(cString.ToString, 4)) = 13 Then
           cString.Remove cString.Length - 2, 2
       End If
    Next i
    
    sTemp = cString.ToString
    
    If iRefactoring = 1 Then
        
        DoReplace sTemp, "0x" & Hex$(Snips.Item(1)), "@start"
        DoReplace sTemp, "0x" & Hex$(Snips.Item(1) + &H8000000), "@start"
        
        If Snips.Count > 1 Then
            
            For i = 2 To Snips.Count
                DoReplace sTemp, "0x" & Hex$(Snips.Item(i)), "@snippet" & i - 1
                DoReplace sTemp, "0x" & Hex$(Snips.Item(i) + &H8000000), "@snippet" & i - 1
            Next i
            
        End If
        
        If Strings.Count > 0 Then
            
            For i = 1 To Strings.Count
                DoReplace sTemp, "0x" & Hex$(Strings.Item(i)), "@string" & i
                DoReplace sTemp, "0x" & Hex$(Strings.Item(i) + &H8000000), "@string" & i
            Next i
            
        End If
        
        If Brailles.Count > 0 Then
        
            For i = 1 To Brailles.Count
                DoReplace sTemp, "0x" & Hex$(Brailles.Item(i)), "@braille" & i
                DoReplace sTemp, "0x" & Hex$(Brailles.Item(i) + &H8000000), "@braille" & i
            Next i
            
        End If
        
        If Moves.Count > 0 Then
            
            For i = 1 To Moves.Count
                DoReplace sTemp, "0x" & Hex$(Moves.Item(i)), "@move" & i
                DoReplace sTemp, "0x" & Hex$(Moves.Item(i) + &H8000000), "@move" & i
            Next i
            
        End If
        
        If Marts.Count > 0 Then
            
            For i = 1 To Marts.Count
                DoReplace sTemp, "0x" & Hex$(Marts.Item(i)), "@mart" & i
                DoReplace sTemp, "0x" & Hex$(Marts.Item(i) + &H8000000), "@mart" & i
            Next i
            
        End If
        
    End If
    
    cString.TheString = sTemp
    
    If iRefactoring = 1 Then
        
        If LenB(sRefactorDynamic) <> 0 Then
            
            If IsHex(sRefactorDynamic) Then
                
                cString.Insert 0, "#dynamic 0x" & sRefactorDynamic & vbNewLine
                
                If cString.Find("#include") = 0 Then
                    cString.Insert 11 + Len(sRefactorDynamic) + 2, vbNewLine
                End If
                
            End If
            
        End If
        
    End If
    
    tmpLong4 = cString.Find("#include")
    
    If tmpLong4 <> 0 Then
    
        Do While cString.Find("#include", tmpLong4) <> 0
            tmpLong4 = tmpLong4 + 8
        Loop
        
        cString.Insert cString.Find(vbNewLine, tmpLong4) + 2 - 1, vbNewLine
        
    End If
    
    Document(frmMain.Tabs.SelectedTab).ClearUndoBuffer
    SendMessageW Document(frmMain.Tabs.SelectedTab).txtCode.hWnd, WM_SETTEXT, 0&, ByVal cString.Pointer
    Document(frmMain.Tabs.SelectedTab).UpdateStack
    IsLoading = False
    
    Document(frmMain.Tabs.SelectedTab).IsDirty = False
    Document(frmMain.Tabs.SelectedTab).Toolbar.BtnState(11) = STA_NORMAL
    frmMain.StatusBar.PanelEnabled(4) = False
    
    UnlockUpdate Document(frmMain.Tabs.SelectedTab).hWnd
    Screen.MousePointer = vbDefault
    
    frmMain.StatusBar.PanelCaption(2) = Document(frmMain.Tabs.SelectedTab).GetCount(Document(frmMain.Tabs.SelectedTab).txtCode)
    SetStatusText LoadResString(13032) & Format$(EndTiming / 1000, "0.000") & LoadResString(13025)
    
    IsLevelScript = Document(frmMain.Tabs.SelectedTab).Toolbar.BtnValue(22)
    
    If IsLevelScript = False Then
        Document(frmMain.Tabs.SelectedTab).Toolbar.BtnState(22) = STA_NORMAL
    Else
        Document(frmMain.Tabs.SelectedTab).Toolbar.BtnState(22) = STA_PRESSED
    End If
  
    '  If Japanese = False Then
    '    Toolbar.BtnState(23) = STA_NORMAL
    '  Else
    '    Toolbar.BtnState(23) = STA_PRESSED
    '  End If
    
    EraseCol Snips
    EraseCol Strings
    EraseCol Moves
    EraseCol Marts
    EraseCol Brailles
    
    Exit Function

FakeLevelScript:
    
    IsLevelScript = False
    Document(frmMain.Tabs.SelectedTab).Toolbar.BtnValue(22) = False
    Document(frmMain.Tabs.SelectedTab).Toolbar.BtnState(22) = STA_NORMAL
    Decompile sFileName, lOffset
    Exit Function
    
ErrHandler:

    Close #iFileNum
    Screen.MousePointer = vbDefault
    UnlockUpdate Document(frmMain.Tabs.SelectedTab).hWnd
    MsgBox LoadResString(13016) & Err.Number & " """ & Err.Description & """.", vbExclamation
    

End Function

