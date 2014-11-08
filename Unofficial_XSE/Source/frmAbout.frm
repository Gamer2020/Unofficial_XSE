VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":000C
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "9000"
   Begin VB.Timer tmrMarquee 
      Interval        =   1500
      Left            =   4200
      Top             =   120
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Yay!"
      Height          =   345
      Left            =   75
      TabIndex        =   0
      Tag             =   "9009"
      Top             =   2685
      Width           =   1095
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "RayGaara4Dragon, liuyanghejerry, Larsie13 - $3"
      ForeColor       =   &H00666666&
      Height          =   195
      Index           =   14
      Left            =   180
      TabIndex        =   16
      Top             =   3600
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   3510
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Linkandzelda, Shadow Lakitu, Christos - $3"
      ForeColor       =   &H00666666&
      Height          =   195
      Index           =   13
      Left            =   180
      TabIndex        =   15
      Top             =   3360
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Brannon, Pascal, Gavin, Malin, SerenadeDS - $3"
      ForeColor       =   &H00666666&
      Height          =   195
      Index           =   12
      Left            =   180
      TabIndex        =   14
      Top             =   3120
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   3450
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "cooley, Matthew, Darthatron, Xiros - $3"
      ForeColor       =   &H00666666&
      Height          =   195
      Index           =   11
      Left            =   180
      TabIndex        =   13
      Top             =   2880
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pool92, =BloodMoon=, $tefano - $3"
      ForeColor       =   &H00666666&
      Height          =   195
      Index           =   10
      Left            =   180
      TabIndex        =   12
      Top             =   2640
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "D-Trogh, Wuggles, ZodiacDaGreat - $5"
      ForeColor       =   &H00666666&
      Height          =   195
      Index           =   9
      Left            =   180
      TabIndex        =   11
      Top             =   2400
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "loadingNOW, Mastermind_X - $4"
      ForeColor       =   &H00666666&
      Height          =   195
      Index           =   8
      Left            =   180
      TabIndex        =   10
      Top             =   2160
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Martin™, TreeckoLv.100 - Beta testing"
      ForeColor       =   &H00666666&
      Height          =   195
      Index           =   7
      Left            =   180
      TabIndex        =   9
      Top             =   1920
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2790
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Green Charizard, Kike-Scott, Liquid_Thunder - $3"
      ForeColor       =   &H00666666&
      Height          =   195
      Index           =   6
      Left            =   180
      TabIndex        =   8
      Top             =   1680
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Serg!o, Zel - $2"
      ForeColor       =   &H00666666&
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   7
      Top             =   1440
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DJ Bouché, Kyoufu Kawa - $1"
      ForeColor       =   &H00666666&
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   6
      Top             =   1200
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Tag             =   "9001"
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   960
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "A whole new scripting experience."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Tag             =   "9002"
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   2460
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Thanks and greetings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Tag             =   "9003"
      Top             =   960
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Label lblMarquee 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2008 HackMew"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   2
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   2025
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   4200
      TabIndex        =   1
      Top             =   2880
      Width           =   390
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lOriginalPos() As Long

' All values here are in milliseconds
Private Const lNormalSpeed = 400
Private Const lFastSpeed = 44
Private Const lDelay = 3300

Private Const lScrollingWidth = 4
Private Const lLabelsToShow = 3

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyCode As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim i As Integer
    
    Localize Me
    
    ' Change the version label with the current values
    lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
    
    ' Resize the array containing the initial top position for the marquee labels
    ReDim lOriginalPos(lblMarquee.UBound) As Long
    
    ' Fill it with the data
    For i = LBound(lOriginalPos) To UBound(lOriginalPos)
        lOriginalPos(i) = lblMarquee(i).Top
    Next i
    
    lblMarquee(2).Caption = App.LegalCopyright
    
    For i = lblMarquee.lBound To lblMarquee.UBound
        If InStrB(lblMarquee(i).Caption, "$1") Then
            lblMarquee(i).Caption = Replace(lblMarquee(i).Caption, "$1", LoadResString(9004))
        ElseIf InStrB(lblMarquee(i).Caption, "$2") Then
            lblMarquee(i).Caption = Replace(lblMarquee(i).Caption, "$2", LoadResString(9005))
        ElseIf InStrB(lblMarquee(i).Caption, "$3") Then
            lblMarquee(i).Caption = Replace(lblMarquee(i).Caption, "$3", LoadResString(9006))
        ElseIf InStrB(lblMarquee(i).Caption, "$4") Then
            lblMarquee(i).Caption = Replace(lblMarquee(i).Caption, "$4", LoadResString(9007))
        ElseIf InStrB(lblMarquee(i).Caption, "$5") Then
            lblMarquee(i).Caption = Replace(lblMarquee(i).Caption, "$5", LoadResString(9008))
        End If
    Next i
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Make scrolling speed faster
    tmrMarquee.Interval = lFastSpeed
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    ' Restore normal scrolling speed
    tmrMarquee.Interval = lNormalSpeed
End Sub

Private Sub tmrMarquee_Timer()
Dim i As Integer
    
    ' If the interval is too high we need to adjust it
    If tmrMarquee.Interval > lNormalSpeed Then
        tmrMarquee.Interval = lNormalSpeed
    End If
    
    ' For each label
    For i = lblMarquee.lBound To lblMarquee.UBound
        
        ' Decrease its top position
        lblMarquee(i).Top = lblMarquee(i).Top - lScrollingWidth
        
        If lblMarquee(i).Top < lOriginalPos(lblMarquee.lBound) - lblMarquee(lblMarquee.lBound).Height \ 2 Then
            ' If the label is too high, hide it
            lblMarquee(i).Visible = False
        ElseIf lblMarquee(i).Top >= lOriginalPos(lLabelsToShow) Then
            ' If the label is too low, hide it
            lblMarquee(i).Visible = False
        Else
            ' Otherwhise make sure it's visible
            lblMarquee(i).Visible = True
        End If
        
    Next i
    
    ' If the last label went too far away
    If lblMarquee(lblMarquee.UBound).Top < -(lblMarquee(lblMarquee.UBound).Top) * 2 Then
        
        For i = lblMarquee.lBound To lblMarquee.UBound
            
            ' Restore initial positions
            lblMarquee(i).Top = lOriginalPos(i)
            
            ' Show the first labels
            If i < lLabelsToShow Then
                lblMarquee(i).Visible = True
            End If
            
        Next i
        
        ' Change the interval so we get a short delay
        tmrMarquee.Interval = lDelay
        
    End If
    
End Sub
