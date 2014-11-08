VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const DEFAULT_CHARSET = 1
Const OUT_DEFAULT_PRECIS = 0
Const CLIP_DEFAULT_PRECIS = 0
Const DEFAULT_QUALITY = 0
Const DEFAULT_PITCH = 0
Const FF_ROMAN = 16
Const CF_PRINTERFONTS = &H2
Const CF_SCREENFONTS = &H1
Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
'Const CF_EFFECTS = &H100&
Const CF_FORCEFONTEXIST = &H10000
Const CF_INITTOLOGFONTSTRUCT = &H40&
Const CF_LIMITSIZE = &H2000&
Const CF_SCALABLEONLY = &H20000
Const CF_NOSCRIPTSEL = &H800000
Const REGULAR_FONTTYPE = &H400
'Const LF_FACESIZE = 32
Const CCHDEVICENAME = 32
Const CCHFORMNAME = 32
Const GMEM_MOVEABLE = &H2
Const GMEM_ZEROINIT = &H40
Const DM_DUPLEX = &H1000&
Const DM_ORIENTATION = &H1&

'Const PD_COLLATE = &H10&
'Const PD_PRINTSETUP = &H40
'Const PD_DISABLEPRINTTOFILE = &H80000
Const PD_HIDEPRINTTOFILE = &H100000
'Const PD_USEDEVMODECOPIES = &H40000


Public Enum PrintFlags
    HidePrintToFile = PD_HIDEPRINTTOFILE
End Enum

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private cdlg As OPENFILENAME
Private LastFileName As String, LastFilter As Long

Private Type OPENFILENAME
    lStructSize       As Long
    hwndOwner         As Long
    hInstance         As Long
    lpsFilter       As String
    lpsCustomFilter As String
    nMaxCustFilter    As Long
    nFilterIndex      As Long
    lpsFile         As String
    nMaxFile          As Long
    lpsFileTitle    As String
    nMaxFileTitle     As Long
    lpsInitialDir   As String
    lpsTitle        As String
    flags             As Long
    nFileOffset       As Integer
    nFileExtension    As Integer
    lpsDefExt       As String
    lCustData         As Long
    lpfnHook          As Long
    lpTemplateName    As String
End Type

Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800

Public Enum DialogFlags
    FileMustExist = OFN_FILEMUSTEXIST
    HideReadOnly = OFN_HIDEREADONLY
    OVERWRITEPROMPT = OFN_OVERWRITEPROMPT
    PATHMUSTEXIST = OFN_PATHMUSTEXIST
End Enum

Private Const sPipe As String = "|"

Private Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Private Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long

Private Type PRINTDLG_TYPE
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type

Private Type DEVMODE_TYPE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'Private Declare Function PAGESETUPDLG Lib "comdlg32.dll" Alias "PageSetupDlgA" (pPagesetupdlg As PAGESETUPDLG) As Long

'Private Type PAGESETUPDLG
'    lStructSize As Long
'    hwndOwner As Long
'    hDevMode As Long
'    hDevNames As Long
'    flags As Long
'    ptPaperSize As POINTAPI
'    rtMinMargin As RECT
'    rtMargin As RECT
'    hInstance As Long
'    lCustData As Long
'    lpfnPageSetupHook As Long
'    lpfnPagePaintHook As Long
'    lpPageSetupTemplateName As String
'    hPageSetupTemplate As Long
'End Type


Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long

'Private Type LOGFONT
'        lfHeight As Long
'        lfWidth As Long
'        lfEscapement As Long
'        lfOrientation As Long
'        lfWeight As Long
'        lfItalic As Byte
'        lfUnderline As Byte
'        lfStrikeOut As Byte
'        lfCharSet As Byte
'        lfOutPrecision As Byte
'        lfClipPrecision As Byte
'        lfQuality As Byte
'        lfPitchAndFamily As Byte
'        lfFaceName As String * 31
'End Type

Private Type CHOOSEFONT
        lStructSize As Long
        hwndOwner As Long          ' caller's window handle
        hDC As Long                ' printer DC/IC or NULL
        lpLogFont As Long          ' ptr. to a LOGFONT suct
        iPointSize As Long         ' 10 * size in points of selected font
        flags As Long              ' enum. type flags
        rgbColors As Long          ' returned text color
        lCustData As Long          ' data passed to hook fn.
        lpfnHook As Long           ' ptr. to hook function
        lpTemplateName As String     ' custom template name
        hInstance As Long          ' instance handle of.EXE that
                                       ' contains cust. dlg. template
        lpszStyle As String          ' return the style field here
                                       ' must be LF_FACESIZE or bigger
        nFontType As Integer          ' same value reported to the EnumFonts
                                       ' call back with the extra FONTTYPE_
                                       ' bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           ' minimum pt size allowed &
        nSizeMax As Long           ' max pt size allowed if
                                       ' CF_LIMITSIZE is used
End Type

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Function ShowOpen(ByVal hWnd As Long, ByVal Title As String, Optional ByVal InitDir As String = "", Optional ByVal Filter As String = "All Files (*.*)|*.*|", Optional flags As DialogFlags = FileMustExist Or HideReadOnly Or PATHMUSTEXIST) As String
    
    DoReplace Filter, sPipe, vbNullChar
    
    If Right$(Filter, 1) <> vbNullChar Then Filter = Filter & vbNullChar
    If LenB(InitDir) = 0 Then InitDir = LastFileName
    
    With cdlg
        .lStructSize = Len(cdlg)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpsFilter = Filter
        .lpsFile = Space$(254)
        .nMaxFile = 255
        .lpsFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpsInitialDir = InitDir
        .lpsTitle = Title
        .flags = flags
    End With
    
    If GetOpenFileName(cdlg) Then
        ShowOpen = Left$(cdlg.lpsFile, InStr(cdlg.lpsFile, vbNullChar) - 1)
        LastFileName = ShowOpen
        LastFilter = cdlg.nFilterIndex
    End If

End Function

Public Function ShowSave(ByVal hWnd As Long, ByVal Title As String, Optional ByVal sFileName As String = "", Optional InitDir As String = "", Optional ByVal Filter As String, Optional flags As DialogFlags = FileMustExist Or OVERWRITEPROMPT Or PATHMUSTEXIST) As String
Dim Buffer As String
Dim Result As Long

    DoReplace Filter, sPipe, vbNullChar
    
    If Right$(Filter, 1) <> vbNullChar Then Filter = Filter & vbNullChar
    Buffer = sFileName & String$(254 - Len(sFileName), 0)
    If LenB(InitDir) = 0 Then InitDir = LastFileName
    
    With cdlg
        .lStructSize = Len(cdlg)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .flags = flags
        .nMaxFile = Len(Buffer)
        .lpsFile = Buffer
        .lpsFilter = Filter
        .lpsInitialDir = InitDir
        .lpsTitle = Title
        .lpsFilter = Filter
        .lpsDefExt = ".rbc"
    End With
    
    Result = GetSaveFileName(cdlg)
    
    If Result <> 0 Then
        ShowSave = Left$(cdlg.lpsFile, InStr(cdlg.lpsFile, vbNullChar) - 1)
        LastFilter = cdlg.nFilterIndex
        LastFileName = ShowSave
    End If

End Function

'Public Property Get FileName() As String
'    FileName = LastFileName
'End Property

Public Property Get FilterIndex() As Long
    FilterIndex = LastFilter
End Property

Public Function ShowColor(ByVal hWnd As Long) As Long
Dim cc As CHOOSECOLOR
    
    'set the sucture size
    cc.lStructSize = Len(cc)
    'Set the owner
    cc.hwndOwner = hWnd
    'set the application's instance
    cc.hInstance = App.hInstance
    'set the custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'no extra flags
    cc.flags = 0

    'Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
    
End Function

Private Sub SetScaleLeft(x As Single, ScaleMode As ScaleModeConstants)
Const PHYSICALOFFSETX As Long = 112

    With Printer
        x = .ScaleX(x, ScaleMode, .ScaleMode)
        .ScaleLeft = .ScaleX(GetDeviceCaps(.hDC, PHYSICALOFFSETX), vbPixels, .ScaleMode) - Abs(x)
    End With
    
End Sub

Private Sub SetScaleTop(Y As Single, ScaleMode As ScaleModeConstants)
Const PHYSICALOFFSETY As Long = 113

    With Printer
        Y = .ScaleY(Y, ScaleMode, .ScaleMode)
        .ScaleTop = .ScaleY(GetDeviceCaps(.hDC, PHYSICALOFFSETY), vbPixels, .ScaleMode) - Abs(Y)
    End With
    
End Sub

Public Sub ShowPrinter(frmOwner As Form, Optional flags As PrintFlags)
Dim PrintDlg As PRINTDLG_TYPE
Dim DevMode As DEVMODE_TYPE
Dim DevName As DEVNAMES_TYPE

Dim lpDevMode As Long, lpDevName As Long
Dim objPrinter As Printer, NewPrinterName As String

Dim sTemp As String
Dim sCaption As String
Dim FontHeight As Single
Dim FontWidth As Single
Dim lMaxLines As Long
Dim lMaxLength As Long
Dim lTotalLines As Long
Dim lTotalPages As Long
Dim sTempArray() As String
Dim sTextArray() As cStringBuilder
Dim i As Long
Dim j As Long

    If HasPrinters = False Then
        Exit Sub
    End If
    
    On Error GoTo Hell
    
    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName suctures
    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hwndOwner = frmOwner.hWnd
    PrintDlg.flags = flags
    
    'Set the current orientation and duplex setting
    DevMode.dmDeviceName = Printer.DeviceName
    DevMode.dmSize = Len(DevMode)
    DevMode.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DevMode.dmPaperWidth = Printer.Width
    DevMode.dmOrientation = Printer.Orientation
    DevMode.dmPaperSize = Printer.PaperSize
    DevMode.dmDuplex = Printer.Duplex

    'Allocate memory for the initialization hDevMode sucture
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevMode))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    
    If lpDevMode > 0 Then
        RtlMoveMemory ByVal lpDevMode, DevMode, Len(DevMode)
        GlobalUnlock PrintDlg.hDevMode
    End If

    'Set the current driver, device, and port name strings
    DevName.wDriverOffset = 8
    DevName.wDeviceOffset = DevName.wDriverOffset + 1 + Len(Printer.DriverName)
    DevName.wOutputOffset = DevName.wDeviceOffset + 1 + Len(Printer.Port)
    DevName.wDefault = 0
    DevName.extra = Printer.DriverName & vbNullChar & Printer.DeviceName & vbNullChar & Printer.Port & vbNullChar

    'Allocate memory for the initial hDevName sucture
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    
    If lpDevName > 0 Then
        RtlMoveMemory ByVal lpDevName, DevName, Len(DevName)
        GlobalUnlock lpDevName
    End If

    'Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) <> 0 Then

        'First get the DevName sucture.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        RtlMoveMemory DevName, ByVal lpDevName, 45
        GlobalUnlock lpDevName
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode sucture and set the printer
        'properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        RtlMoveMemory DevMode, ByVal lpDevMode, Len(DevMode)
        GlobalUnlock PrintDlg.hDevMode
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left$(DevMode.dmDeviceName, InStr(DevMode.dmDeviceName, vbNullChar) - 1))
        
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                End If
            Next
        End If

        'Set printer object properties according to selections made by user
        Printer.Copies = DevMode.dmCopies
        Printer.Duplex = DevMode.dmDuplex
        Printer.Orientation = DevMode.dmOrientation
        Printer.PaperSize = DevMode.dmPaperSize
        Printer.PrintQuality = DevMode.dmPrintQuality
        Printer.ColorMode = DevMode.dmColor
        Printer.PaperBin = DevMode.dmDefaultSource
        
        Printer.Print vbNullString 'Initialize printer
        
        Printer.FontName = frmOwner.txtCode.FontName
        Printer.FontSize = frmOwner.txtCode.FontSize
        Printer.FontBold = frmOwner.txtCode.FontBold
        Printer.FontItalic = frmOwner.txtCode.FontItalic
        Printer.ForeColor = frmOwner.txtCode.ForeColor
        Printer.FontTransparent = True
        
        sCaption = frmOwner.Caption
        Printer.ScaleMode = vbMillimeters
        SetScaleLeft -20, vbMillimeters
        SetScaleTop -25, vbMillimeters
        
        FontWidth = Printer.TextWidth(vbSpace)
        FontHeight = Printer.TextHeight(vbSpace)
        lMaxLines = Int((Printer.ScaleHeight - Abs(Printer.ScaleTop) * 2) / FontHeight) - 4
        lMaxLength = Int((Printer.ScaleWidth - Abs(Printer.ScaleLeft) * 2) / FontWidth)
        
        SplitB frmOwner.txtCode.text, sTempArray, vbNewLine
               
        For i = LBound(sTempArray) To UBound(sTempArray)
            If Len(sTempArray(i)) > lMaxLength Then
                sTemp = Left$(sTempArray(i), lMaxLength)
                For j = 1 To (Len(sTempArray(i)) - (Len(sTempArray(i)) Mod lMaxLength)) \ lMaxLength
                    sTemp = sTemp & vbNewLine & Mid$(sTempArray(i), (lMaxLength * j) + 1, lMaxLength)
                Next j
                sTempArray(i) = sTemp
            End If
        Next i
        
        sTemp = Join(sTempArray, vbNewLine)
        SplitB sTemp, sTempArray, vbNewLine
                
        lTotalLines = UBound(sTempArray) + 1
        lTotalPages = lTotalLines \ lMaxLines
        ReDim sTextArray(lTotalPages) As cStringBuilder
        
        'SetScaleLeft 0, vbMillimeters
        Printer.ScaleLeft = 0
        Printer.CurrentX = Int(Printer.ScaleWidth - Printer.TextWidth(sCaption)) \ 2
        Printer.CurrentY = 0
        Printer.Print sCaption & vbNewLine
        
        SetScaleLeft -20, vbMillimeters
        Printer.CurrentX = 0
        Printer.CurrentY = FontHeight * 2
        
        For i = 0 To lTotalPages - 1
            Set sTextArray(i) = New cStringBuilder
            For j = 0 To lMaxLines - 1
                sTextArray(i).Append sTempArray(j + lMaxLines * i) & vbNewLine
            Next j
            If (UBound(sTempArray) - lMaxLines * (i + 1)) < 0 Then
                sTextArray(i).Append String$(-(UBound(sTempArray) - lMaxLines * (i + 1)) - 1, vbNewLine)
            End If
        Next i
        
        Erase sTempArray
        
        For i = 0 To lTotalPages - 1
            
            Printer.Print sTextArray(i).ToString
            Printer.Print vbNullString
            
            Printer.ScaleLeft = 0
            Printer.CurrentX = Int(Printer.ScaleWidth - Printer.TextWidth(CStr(i + 1))) \ 2
            Printer.Print i + 1
            
            If i > 0 Then
                If i < lTotalPages Then
                    
                    Printer.NewPage
                    Printer.CurrentX = Int(Printer.ScaleWidth - Printer.TextWidth(sCaption)) \ 2
                    Printer.CurrentY = 0
                    Printer.Print sCaption & vbNewLine
                    
                    SetScaleLeft -20, vbMillimeters
                    Printer.CurrentX = 0
                    Printer.CurrentY = FontHeight * 2
                    
                End If
            End If
            
        Next i
        
        Printer.EndDoc
        SetStatusText LoadResString(13041)
        
    End If
    
    GlobalFree PrintDlg.hDevMode
    GlobalFree PrintDlg.hDevNames
    
Hell:
End Sub

'Public Function ShowPageSetupDlg(ByVal hWnd As Long) As Long
'    Dim m_PSD As PAGESETUPDLG
'    'Set the sucture size
'    m_PSD.lStructSize = Len(m_PSD)
'    'Set the owner window
'    m_PSD.hwndOwner = hWnd
'    'Set the application instance
'    m_PSD.hInstance = App.hInstance
'    'no extra flags
'    m_PSD.flags = 0
'
'    'Show the pagesetup dialog
'    If PAGESETUPDLG(m_PSD) Then
'        ShowPageSetupDlg = 0
'    Else
'        ShowPageSetupDlg = -1
'    End If
'End Function

Public Function ShowFont(frmOwner As Form) As String
Dim cf As CHOOSEFONT, lfont As LOGFONT, hMem As Long, pMem As Long
Dim RetVal As Long ',fontname As String

    lfont.lfHeight = Int(frmOwner.txtCode.FontSize * (4 / 3)) ' determine default height (pt + 3)
    lfont.lfWidth = 0  ' determine default width
    lfont.lfEscapement = 0  ' angle between baseline and escapement vector
    lfont.lfOrientation = 0  ' angle between baseline and orientation vector
    
    If frmOwner.txtCode.FontBold Then
        lfont.lfWeight = FW_BOLD  ' normal weight i.e. not bold
    Else
        lfont.lfWeight = FW_NORMAL
    End If
    
    lfont.lfItalic = frmOwner.txtCode.FontItalic
    lfont.lfCharSet = DEFAULT_CHARSET  ' use default character set
    lfont.lfOutPrecision = OUT_DEFAULT_PRECIS  ' default precision mapping
    lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS  ' default clipping precision
    lfont.lfQuality = DEFAULT_QUALITY  ' default quality setting
    lfont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN  ' default pitch, proportional with serifs
    lfont.lfFaceName = frmOwner.txtCode.font & vbNullChar  ' string must be null-terminated
    
    ' Create the memory block which will act as the LOGFONT sucture buffer.
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)  ' lock and get pointer
    RtlMoveMemory ByVal pMem, lfont, Len(lfont)  ' copy sucture's contents into block
    
    ' Initialize dialog box: Screen and printer fonts, point size between 8 and 72.
    cf.lStructSize = Len(cf)  ' size of sucture
    cf.hwndOwner = frmOwner.hWnd
'    cf.hdc = Printer.hdc  ' device context of default printer
    cf.lpLogFont = pMem   ' pointer to LOGFONT memory block buffer
    cf.iPointSize = 90   ' 9 point font (in units of 1/10 point)
    cf.flags = CF_BOTH Or CF_SCALABLEONLY Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE Or CF_NOSCRIPTSEL
    cf.rgbColors = vbBlack
    cf.nFontType = REGULAR_FONTTYPE  ' regular font type i.e. not bold or anything
    cf.nSizeMin = 8  ' minimum point size
    cf.nSizeMax = 72  ' maximum point size
    
    ' Now, call the function. If successful, copy the LOGFONT sucture back into the sucture
    ' and then print out the attributes we mentioned earlier that the user selected.
    RetVal = CHOOSEFONT(cf)  ' open the dialog box
    
    If RetVal <> 0 Then  ' success
        RtlMoveMemory lfont, ByVal pMem, Len(lfont)  ' copy memory back
        ' Now make the fixed-length string holding the font name into a "normal" string.
        ShowFont = Left$(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
        frmOwner.txtCode.FontSize = CSng(cf.iPointSize / 10)
        If lfont.lfWeight = FW_NORMAL Then
            frmOwner.txtCode.FontBold = False
        Else
            frmOwner.txtCode.FontBold = True
        End If
        frmOwner.txtCode.FontItalic = lfont.lfItalic
    End If
    
    ' Deallocate the memory block we created earlier. Note that this must
    ' be done whether the function succeeded or not.
    RetVal = GlobalUnlock(hMem)  ' destroy pointer, unlock block
    RetVal = GlobalFree(hMem)  ' free the allocated memory
    
End Function
