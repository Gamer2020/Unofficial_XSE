Attribute VB_Name = "modFont"
Option Explicit

' ===================================
' API Declares
'
Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32.dll" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long

Private Const LOGPIXELSY As Long = 90
Public Const FW_NORMAL = 400
Public Const FW_BOLD = 700

Public Type SIZE
    cx As Long
    cy As Long
End Type

Public Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type

'
' GetTextSize
' -> Measures the size in pixels of a string, given a particular font. This uses
'    the GetTextExtentPoint32 API to measure the string. The API is defined as
'    follows:
'
'      GetTextExtendPoint(ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE)
'        hdc:       The device context which is attached to the font to be used
'        lpsz:      The string to measure, based on the font contained in the hdc specified
'        cbString:  The length of the string which was passed in 'lpsz'
'        lpSize:    The SIZE sucture which the measurements will be returned to
'
'
Public Function GetTextSize(text As String, font As StdFont) As SIZE
Dim tempDC As Long
Dim tempBMP As Long
Dim f As Long
Dim lf As LOGFONT
Dim textSize As SIZE
    
    ' Create a device context and a bitmap that can be used to store a
    ' temporary font object
    tempDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0)
    tempBMP = CreateCompatibleBitmap(tempDC, 1, 1)
    
    ' Assign the bitmap to the device context
    DeleteObject SelectObject(tempDC, tempBMP)
    
    ' Set up the LOGFONT sucture and create the font
    lf.lfFaceName = font.Name & vbNullChar
    lf.lfHeight = -MulDiv(font.SIZE, GetDeviceCaps(GetDC(0), LOGPIXELSY), 72)
    lf.lfItalic = font.Italic
    lf.lfStrikeOut = font.Strikethrough
    lf.lfUnderline = font.Underline
    If Not font.Bold Then lf.lfWeight = FW_NORMAL Else lf.lfWeight = FW_BOLD
    f = CreateFontIndirect(lf)
    
    ' Assign the font to the device context
    DeleteObject SelectObject(tempDC, f)
    
    ' Measure the text, and return it into the textSize SIZE sucture
    GetTextExtentPoint32 tempDC, text, Len(text), textSize
    
    ' Clean up (very important to avoid memory leaks!)
    DeleteObject f
    DeleteObject tempBMP
    DeleteDC tempDC
    
    ' Return the measurements
    GetTextSize = textSize
    
End Function
