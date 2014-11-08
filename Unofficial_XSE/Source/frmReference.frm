VERSION 5.00
Begin VB.Form frmReference 
   BackColor       =   &H00D0D0D0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Command Help"
   ClientHeight    =   1815
   ClientLeft      =   6765
   ClientTop       =   7395
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000011&
   Icon            =   "frmReference.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   121
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "12000"
   Begin VB.ComboBox cboList 
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Text            =   "<cmd>"
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblNeededBytes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<needed bytes>"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblParams 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<params>"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00B6B6B6&
      X1              =   8
      X2              =   352
      Y1              =   64
      Y2              =   64
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<description>"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<cmd>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00EAEAEA&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00B6B6B6&
      Height          =   1575
      Left            =   120
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsDeleting As Boolean

Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_ADDSTRING = &H143

Private Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Private Function GetSizeName(bSize As Byte) As String
    Select Case bSize
        Case 0: GetSizeName = "Byte"
        Case 1: GetSizeName = "Word"
        Case 2: GetSizeName = "DWord"
        Case 3: GetSizeName = "Pointer"
        Case Else: GetSizeName = "???"
    End Select
End Function

Public Sub ResizeMe()
    Shape1.Width = ScaleWidth - 16
    Line2.Y1 = lblDescription.Top + lblDescription.Height + 8
    Line2.Y2 = Line2.Y1
    Line2.X2 = Shape1.Width + 8
    cboList.Left = ScaleWidth - 16 - cboList.Width
    lblCommand.Width = Shape1.Width - 16
    lblDescription.Width = Shape1.Width - 16
    lblNeededBytes.Top = Line2.Y1 + 8
    lblParams.Width = Shape1.Width - 16
    lblParams.Top = lblNeededBytes.Top + 16
    Shape1.Height = lblParams.Top + lblParams.Height + 8
    Height = (Shape1.Height + 40) * Screen.TwipsPerPixelY
End Sub

Private Sub cboList_Change()
Dim lRet As Long
Dim lSelStart As Long
    
    If IsDeleting Then
        Exit Sub
    End If
    
    lRet = SendMessageStr(cboList.hWnd, CB_FINDSTRING, -1, cboList.text)
    lSelStart = cboList.SelStart
    
    If lRet > -1 Then
        cboList.ListIndex = lRet
        cboList.SelStart = lSelStart
        cboList.SelLength = Len(cboList.text)
    End If
    
End Sub

Private Sub cboList_Click()
Dim i As Integer, j As Integer
Dim sList As String

    sList = cboList.List(cboList.ListIndex)
    
    If cboList.ListIndex < &HE3 Then
             
        For i = LBound(RubiCommands) To UBound(RubiCommands)
            If LenB(sList) = LenB(RubiCommands(i).Keyword) Then
                If sList = RubiCommands(i).Keyword Then
                  lblCommand.Caption = "0x" & Right$("0" & Hex$(i), 2) & " - " & RubiCommands(i).Keyword
                  lblDescription.Caption = RubiCommands(i).Description
                  lblNeededBytes.Caption = LoadResString(12001) & RubiCommands(i).NeededBytes
                  If RubiCommands(i).ParamCount = 0 Then
                    lblParams.Caption = LoadResString(12002)
                  Else
                    lblParams.Caption = LoadResString(12003)
                    For j = 0 To RubiCommands(i).ParamCount - 1
                      lblParams.Caption = frmReference.lblParams & vbNewLine & _
                                " › " & GetSizeName(RubiParams(i, j).SIZE) & " - " & _
                                RubiParams(i, j).Description
                    Next j
                  End If
                  Exit For
                End If
            End If
        Next i
  
    ElseIf cboList.ListIndex > &HE3 Then
        
        lblNeededBytes.Caption = LoadResString(12001)
        lblParams.Caption = LoadResString(12003) & vbNewLine
        
        Select Case cboList.ListIndex
            Case &HE4
                lblCommand.Caption = "msgbox, message"
                lblDescription.Caption = "Loads a pointer into memory to display a message later on."
                lblNeededBytes.Caption = lblNeededBytes.Caption & 8
                lblParams.Caption = lblParams.Caption & " › " & GetSizeName(RubiParams(&HF, 1).SIZE) & " - " & _
                                    RubiParams(&HF, 1).Description & vbNewLine & _
                                    " › " & GetSizeName(0) & " - Message type"
            Case &HE5
                lblCommand.Caption = sList 'giveitem
                lblDescription.Caption = "Gives a specified item and displays an aftermath message of the player receiving the item."
                lblNeededBytes.Caption = lblNeededBytes.Caption & 12
                lblParams.Caption = lblParams.Caption & " › " & GetSizeName(RubiParams(&H44, 0).SIZE) & " - " & _
                                    RubiParams(&H44, 0).Description & vbNewLine & _
                                    " › " & GetSizeName(RubiParams(&H44, 1).SIZE) & " - " & _
                                    RubiParams(&H44, 1).Description & vbNewLine & _
                                    " › " & GetSizeName(0) & " - Message type"
            Case &HE6
                lblCommand.Caption = sList 'giveitem2
                lblDescription.Caption = "Similar to giveitem except it plays a fanfare too."
                lblNeededBytes.Caption = lblNeededBytes.Caption & 17
                lblParams.Caption = lblParams.Caption & " › " & GetSizeName(RubiParams(&H44, 0).SIZE) & " - " & _
                                    RubiParams(&H44, 0).Description & vbNewLine & _
                                    " › " & GetSizeName(RubiParams(&H44, 1).SIZE) & " - " & _
                                    RubiParams(&H44, 1).Description & vbNewLine & _
                                    " › " & GetSizeName(RubiParams(&H31, 0).SIZE) & " - " & _
                                    RubiParams(&H31, 0).Description
            Case &HE7
                lblCommand.Caption = sList 'giveitem3
                lblDescription.Caption = "Gives the player a specified decoration and displays a related message."
                lblNeededBytes.Caption = lblNeededBytes.Caption & 7
                lblParams.Caption = lblParams.Caption & " › " & GetSizeName(RubiParams(&H4B, 0).SIZE) & " - " & _
                                    RubiParams(&H4B, 0).Description
            Case &HE8
                lblCommand.Caption = sList 'wildbattle
                lblDescription.Caption = "Starts a wild Pokémon battle."
                lblNeededBytes.Caption = lblNeededBytes.Caption & 7
                lblParams.Caption = lblParams.Caption & " › " & GetSizeName(RubiParams(&H79, 0).SIZE) & " - " & _
                                    "Pokémon species to battle" & vbNewLine & _
                                    " › " & GetSizeName(RubiParams(&H79, 1).SIZE) & " - " & _
                                    RubiParams(&H79, 1).Description & vbNewLine & _
                                    " › " & GetSizeName(RubiParams(&H79, 2).SIZE) & " - " & _
                                    RubiParams(&H79, 2).Description
            Case &HE9
                lblCommand.Caption = sList 'wildbattle2
                lblDescription.Caption = "Starts a wild battle using a specific graphic style."
                lblNeededBytes.Caption = lblNeededBytes.Caption & 10
                lblParams.Caption = lblParams.Caption & " › " & GetSizeName(RubiParams(&H79, 0).SIZE) & " - " & _
                                    "Pokémon species to battle" & vbNewLine & _
                                    " › " & GetSizeName(RubiParams(&H79, 1).SIZE) & " - " & _
                                    RubiParams(&H79, 1).Description & vbNewLine & _
                                    " › " & GetSizeName(RubiParams(&H79, 2).SIZE) & " - " & _
                                    RubiParams(&H79, 2).Description & vbNewLine & _
                                    " › " & GetSizeName(0) & " - Battle style"
            Case &HEA
                lblCommand.Caption = sList 'registernav
                lblDescription.Caption = "Register the specified trainer in the PokéNav. Emerald only."
                lblNeededBytes.Caption = lblNeededBytes.Caption & 7
                lblParams.Caption = lblParams.Caption & " › " & GetSizeName(1) & " - Trainer ID #"

        End Select
        
    End If
    
    ResizeMe

End Sub

Private Sub cboList_GotFocus()
    cboList.SelStart = 0
    cboList.SelLength = Len(cboList.text)
End Sub

Private Sub cboList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lRet As Long
    
    Select Case KeyCode
        
        Case vbKeyDelete, vbKeyBack
            IsDeleting = True
            
        Case vbKeyReturn
        
            IsDeleting = False
            KeyCode = 0
            
            lRet = SendMessageStr(cboList.hWnd, CB_FINDSTRINGEXACT, -1, Left$(cboList.text, Len(cboList.text) - cboList.SelLength))
        
            If lRet > -1 Then
                cboList.ListIndex = lRet
            End If
            
            cboList.SelStart = Len(cboList.text)
            cboList.SelLength = 0
            
        Case Else
            IsDeleting = False
            
    End Select
    
End Sub

Private Sub cboList_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyTab Then
        KeyCode = 0
    End If
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
    
    Localize Me
    
    LockUpdate cboList.hWnd
    
    For i = LBound(RubiCommands) To UBound(RubiCommands)
        SendMessageStr cboList.hWnd, CB_ADDSTRING, 0, ByVal RubiCommands(i).Keyword
    Next i
    
    SendMessageStr cboList.hWnd, CB_ADDSTRING, 0, ByVal "--------"
    SendMessageStr cboList.hWnd, CB_ADDSTRING, 0, ByVal "msgbox"
    SendMessageStr cboList.hWnd, CB_ADDSTRING, 0, ByVal "giveitem"
    SendMessageStr cboList.hWnd, CB_ADDSTRING, 0, ByVal "giveitem2"
    SendMessageStr cboList.hWnd, CB_ADDSTRING, 0, ByVal "giveitem3"
    SendMessageStr cboList.hWnd, CB_ADDSTRING, 0, ByVal "wildbattle"
    SendMessageStr cboList.hWnd, CB_ADDSTRING, 0, ByVal "wildbattle2"
    SendMessageStr cboList.hWnd, CB_ADDSTRING, 0, ByVal "registernav"
    
    UnlockUpdate cboList.hWnd
    
End Sub
