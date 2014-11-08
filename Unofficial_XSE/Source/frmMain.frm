VERSION 5.00
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000F&
   Caption         =   "XSE - eXtreme Script Editor"
   ClientHeight    =   6885
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11625
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "System"
   LockControls    =   -1  'True
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrKeyboard 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   8040
      Top             =   6120
   End
   Begin VB.Timer tmrRestore 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   8520
      Top             =   6120
   End
   Begin VB.PictureBox picSidebar 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6075
      Left            =   9015
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   174
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   480
      Width           =   2610
      Begin VB.CommandButton cmdOfs 
         Caption         =   "ofs"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2085
         TabIndex        =   34
         Top             =   2550
         Width           =   420
      End
      Begin VB.CommandButton cmdPtr 
         Caption         =   "ptr"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1605
         TabIndex        =   33
         Top             =   2550
         Width           =   420
      End
      Begin VB.CommandButton cmdMR 
         Caption         =   "MR"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   645
         TabIndex        =   36
         Top             =   2895
         Width           =   420
      End
      Begin VB.CommandButton cmdMS 
         Caption         =   "MS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         TabIndex        =   35
         Top             =   2895
         Width           =   420
      End
      Begin VB.TextBox txtMem 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   105
         Width           =   360
      End
      Begin VB.OptionButton optDecHex 
         Caption         =   "Hex"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   195
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton cmdCE 
         Caption         =   "CE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1605
         TabIndex        =   3
         Top             =   480
         Width           =   420
      End
      Begin VB.TextBox txtDisplay 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   615
         Locked          =   -1  'True
         MaxLength       =   11
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   105
         Width           =   1890
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2085
         TabIndex        =   4
         Top             =   480
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "E"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   1125
         TabIndex        =   7
         Top             =   825
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "F"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   1605
         TabIndex        =   8
         Top             =   825
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "C"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   165
         TabIndex        =   5
         Top             =   825
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "D"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   645
         TabIndex        =   6
         Top             =   825
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   645
         TabIndex        =   11
         Top             =   1170
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   165
         TabIndex        =   10
         Top             =   1170
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "B"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   1605
         TabIndex        =   13
         Top             =   1170
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "A"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   1125
         TabIndex        =   12
         Top             =   1170
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   645
         TabIndex        =   16
         Top             =   1515
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   165
         TabIndex        =   15
         Top             =   1515
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   1605
         TabIndex        =   18
         Top             =   1515
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   1125
         TabIndex        =   17
         Top             =   1515
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   645
         TabIndex        =   21
         Top             =   1860
         Width           =   420
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2085
         TabIndex        =   24
         Top             =   1860
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1605
         TabIndex        =   23
         Top             =   1860
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1125
         TabIndex        =   22
         Top             =   1860
         Width           =   420
      End
      Begin VB.CommandButton cmdMultiply 
         Caption         =   "×"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2085
         TabIndex        =   14
         Top             =   1170
         Width           =   420
      End
      Begin VB.CommandButton cmdSubtract 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2085
         TabIndex        =   19
         Top             =   1515
         Width           =   420
      End
      Begin VB.CommandButton cmdNum 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   165
         TabIndex        =   20
         Top             =   1860
         Width           =   420
      End
      Begin VB.CommandButton cmdDivide 
         Caption         =   "÷"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2085
         TabIndex        =   9
         Top             =   825
         Width           =   420
      End
      Begin VB.CommandButton cmdPlusMinus 
         Caption         =   "+/-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1605
         TabIndex        =   28
         Top             =   2205
         Width           =   420
      End
      Begin VB.CommandButton cmdEqual 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2085
         TabIndex        =   29
         Top             =   2205
         Width           =   420
      End
      Begin VB.CommandButton cmdOr 
         Caption         =   "or"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   645
         TabIndex        =   26
         Top             =   2205
         Width           =   420
      End
      Begin VB.CommandButton cmdAnd 
         Caption         =   "and"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         TabIndex        =   25
         Top             =   2205
         Width           =   420
      End
      Begin VB.CommandButton cmdLs 
         Caption         =   "lsh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   645
         TabIndex        =   31
         Top             =   2550
         Width           =   420
      End
      Begin VB.CommandButton cmdNot 
         Caption         =   "not"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   165
         TabIndex        =   30
         Top             =   2550
         Width           =   420
      End
      Begin VB.CommandButton cmdXor 
         Caption         =   "xor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1125
         TabIndex        =   27
         Top             =   2205
         Width           =   420
      End
      Begin VB.CommandButton cmdRs 
         Caption         =   "rsh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1125
         TabIndex        =   32
         Top             =   2550
         Width           =   420
      End
      Begin VB.OptionButton optDecHex 
         Caption         =   "Dec"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   885
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtCommandLine 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1740
         LinkItem        =   "txtCommandLine"
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.TextBox txtNotes 
         BackColor       =   &H00B7F9FE&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   165
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3330
         Width           =   2340
      End
      Begin VB.Image imgCrossOff 
         Height          =   135
         Left            =   480
         Picture         =   "frmMain.frx":000C
         Top             =   5760
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imgDownOff 
         Height          =   60
         Left            =   720
         Picture         =   "frmMain.frx":005D
         Top             =   5760
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Image imgLeftOff 
         Height          =   120
         Left            =   960
         Picture         =   "frmMain.frx":00A2
         Top             =   5760
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgRightOff 
         Height          =   120
         Left            =   1080
         Picture         =   "frmMain.frx":00E7
         Top             =   5760
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgCrossOn 
         Height          =   135
         Left            =   1320
         Picture         =   "frmMain.frx":012D
         Top             =   5760
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imgDownOn 
         Height          =   60
         Left            =   1560
         Picture         =   "frmMain.frx":017E
         Top             =   5760
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Image imgLeftOn 
         Height          =   120
         Left            =   1800
         Picture         =   "frmMain.frx":01C3
         Top             =   5760
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Image imgRightOn 
         Height          =   120
         Left            =   1920
         Picture         =   "frmMain.frx":0208
         Top             =   5760
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Line linSidebar 
         BorderColor     =   &H00C0C0C0&
         X1              =   2
         X2              =   2
         Y1              =   5
         Y2              =   404
      End
   End
   Begin VB.PictureBox picTabs 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   11625
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   11625
      Begin VB.PictureBox picTabControl 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10515
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   35
         Width           =   970
         Begin VB.Image imgClose 
            Height          =   135
            Left            =   780
            Top             =   60
            Width           =   135
         End
         Begin VB.Image imgTabs 
            Height          =   60
            Left            =   555
            Top             =   120
            Width           =   120
         End
         Begin VB.Image imgNext 
            Height          =   120
            Left            =   345
            Top             =   75
            Width           =   60
         End
         Begin VB.Image imgPrev 
            Height          =   120
            Left            =   90
            Top             =   75
            Width           =   60
         End
         Begin VB.Label lblTabs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   510
            TabIndex        =   48
            Top             =   0
            Width           =   225
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00C8C8C8&
            X1              =   48
            X2              =   48
            Y1              =   0
            Y2              =   16
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C8C8C8&
            X1              =   32
            X2              =   32
            Y1              =   0
            Y2              =   16
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00C8C8C8&
            X1              =   16
            X2              =   16
            Y1              =   0
            Y2              =   16
         End
         Begin VB.Label lblPrev 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   30
            TabIndex        =   41
            Top             =   0
            Width           =   225
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C8C8C8&
            Height          =   255
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lblClose 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Wingdings 2"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   735
            TabIndex        =   42
            Top             =   15
            Width           =   225
         End
         Begin VB.Label lblNext 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Webdings"
               Size            =   8.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   270
            TabIndex        =   40
            Top             =   0
            Width           =   195
         End
      End
      Begin eXtremeScriptEditor.TabControl Tabs 
         Height          =   465
         Left            =   60
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   30
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picStatusBar 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   775
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   6555
      Width           =   11625
      Begin eXtremeScriptEditor.xpWellsStatusBar StatusBar 
         Height          =   330
         Left            =   0
         Top             =   0
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   582
         BackColor       =   16053492
         ForeColor       =   0
         ForeColorDissabled=   9474192
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumberOfPanels  =   7
         MaskColor       =   0
         PWidth1         =   249
         pText1          =   ""
         pTTText1        =   ""
         pEnabled1       =   -1  'True
         PWidth2         =   178
         pText2          =   ""
         pTTText2        =   ""
         pEnabled2       =   -1  'True
         PWidth3         =   176
         pText3          =   "Copyright © 2008 HackMew"
         pTTText3        =   ""
         pEnabled3       =   0   'False
         PWidth4         =   18
         pText4          =   "*"
         pTTText4        =   ""
         pEnabled4       =   0   'False
         PWidth5         =   37
         pText5          =   "CAPS"
         pTTText5        =   ""
         pEnabled5       =   0   'False
         PWidth6         =   35
         pText6          =   "NUM"
         pTTText6        =   ""
         pEnabled6       =   0   'False
         PWidth7         =   37
         pText7          =   "SCRL"
         pTTText7        =   ""
         pEnabled7       =   0   'False
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      HelpContextID   =   1
      Begin VB.Menu mnuNew 
         Caption         =   "New Tab"
         HelpContextID   =   2
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open..."
         HelpContextID   =   3
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         HelpContextID   =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As..."
         Enabled         =   0   'False
         HelpContextID   =   5
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
         Enabled         =   0   'False
         HelpContextID   =   63
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
         Enabled         =   0   'False
         HelpContextID   =   6
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecentFiles 
         Caption         =   "Recent Files"
         Enabled         =   0   'False
         HelpContextID   =   7
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   0
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRecent 
            Caption         =   ""
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSep11 
            Caption         =   "-"
         End
         Begin VB.Menu mnuClearRecent 
            Caption         =   "Clear"
            HelpContextID   =   57
         End
      End
      Begin VB.Menu mnuSep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         HelpContextID   =   8
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      HelpContextID   =   9
      Begin VB.Menu mnuUndo 
         Caption         =   "Undo"
         Enabled         =   0   'False
         HelpContextID   =   10
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "Redo"
         Enabled         =   0   'False
         HelpContextID   =   60
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
         HelpContextID   =   11
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Enabled         =   0   'False
         HelpContextID   =   12
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
         HelpContextID   =   13
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         HelpContextID   =   14
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuRevert 
         Caption         =   "Revert"
         Enabled         =   0   'False
         HelpContextID   =   15
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Backup..."
         Enabled         =   0   'False
         HelpContextID   =   16
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuReadOnly 
         Caption         =   "Read Only"
         HelpContextID   =   17
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find and Replace..."
         Enabled         =   0   'False
         HelpContextID   =   18
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuGoto 
         Caption         =   "Goto..."
         HelpContextID   =   19
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
         HelpContextID   =   20
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDateTime 
         Caption         =   "Insert Date/Time"
         HelpContextID   =   21
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "Format"
      HelpContextID   =   22
      Begin VB.Menu mnuBackgroundColor 
         Caption         =   "Background Color"
         HelpContextID   =   23
      End
      Begin VB.Menu mnuForegroundColor 
         Caption         =   "Foreground Color"
         HelpContextID   =   24
      End
      Begin VB.Menu mnuResetColors 
         Caption         =   "Reset Colors"
         HelpContextID   =   25
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font..."
         HelpContextID   =   26
      End
      Begin VB.Menu mnuFontSize 
         Caption         =   "Font Size"
         HelpContextID   =   27
         Begin VB.Menu mnuIncrease 
            Caption         =   "Zoom In"
            HelpContextID   =   29
         End
         Begin VB.Menu mnuDecrease 
            Caption         =   "Zoom Out"
            HelpContextID   =   28
         End
      End
      Begin VB.Menu mnuResetFont 
         Caption         =   "Reset Font"
         HelpContextID   =   30
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      HelpContextID   =   31
      Begin VB.Menu mnuAlwaysonTop 
         Caption         =   "Always on Top"
         HelpContextID   =   32
      End
      Begin VB.Menu mnuMinimizetoSystemTray 
         Caption         =   "Minimize to Sytem Tray"
         HelpContextID   =   33
      End
      Begin VB.Menu mnuRememberSize 
         Caption         =   "Remember Window Size"
         HelpContextID   =   62
      End
      Begin VB.Menu mnu4b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoSave 
         Caption         =   "Auto Save"
         HelpContextID   =   17000
      End
      Begin VB.Menu mnuIgnoreChanges 
         Caption         =   "Ignore Changes on Close"
         HelpContextID   =   35
      End
      Begin VB.Menu mnuShowRecentFiles 
         Caption         =   "Show Recent Files"
         Checked         =   -1  'True
         HelpContextID   =   36
      End
      Begin VB.Menu mnuSep4a 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLineNumbers 
         Caption         =   "Line Numbers"
         Checked         =   -1  'True
         HelpContextID   =   34
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuInlineCommandHelp 
         Caption         =   "Inline Command Help"
         Checked         =   -1  'True
         HelpContextID   =   68
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuBar 
         Caption         =   "Menu Bar"
         Checked         =   -1  'True
         HelpContextID   =   47
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuStatusBar 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
         HelpContextID   =   37
      End
      Begin VB.Menu mnuSidebar 
         Caption         =   "Sidebar"
         Checked         =   -1  'True
         HelpContextID   =   38
      End
      Begin VB.Menu mnuSidebarAlignment 
         Caption         =   "Sidebar Alignment"
         HelpContextID   =   39
         Begin VB.Menu mnuRight 
            Caption         =   "Right"
            Checked         =   -1  'True
            HelpContextID   =   40
         End
         Begin VB.Menu mnuLeft 
            Caption         =   "Left"
            HelpContextID   =   41
         End
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAssociate 
         Caption         =   "Associate..."
         HelpContextID   =   42
      End
      Begin VB.Menu mnuDecompileOptions 
         Caption         =   "Decompile Options"
         HelpContextID   =   64
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      HelpContextID   =   43
      Begin VB.Menu mnuBuiltinScripts 
         Caption         =   "Script Templates"
         HelpContextID   =   45
      End
      Begin VB.Menu mnuBatchCompiler 
         Caption         =   "Batch Compiler"
         HelpContextID   =   44
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuTextAdjuster 
         Caption         =   "Text Adjuster"
         HelpContextID   =   46
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHexViewer 
         Caption         =   "Hex Viewer"
         Enabled         =   0   'False
         HelpContextID   =   14000
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuExpander 
         Caption         =   "ROM Resizer"
         Enabled         =   0   'False
         HelpContextID   =   65
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep15b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFreeSpaceFinder 
         Caption         =   "Free Space Finder"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAdvanceTrainer 
         Caption         =   "Advance Trainer"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      HelpContextID   =   48
      Begin VB.Menu mnuCommandHelp 
         Caption         =   "Command Help"
         HelpContextID   =   61
      End
      Begin VB.Menu mnuGuide 
         Caption         =   "Guide"
         HelpContextID   =   49
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         HelpContextID   =   50
      End
   End
   Begin VB.Menu mnuTabs 
      Caption         =   "Tabs"
      Visible         =   0   'False
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   0
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   1
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   2
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   3
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   4
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   5
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   6
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   7
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   8
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   9
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   10
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   11
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   12
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   13
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   14
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   15
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   16
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   17
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   18
      End
      Begin VB.Menu mnuTab 
         Caption         =   "mnuTab"
         Index           =   19
      End
   End
   Begin VB.Menu mnuSidebarPopup 
      Caption         =   "SidebarPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuHideSidebar 
         Caption         =   "Hide Sidebar"
         HelpContextID   =   52
      End
      Begin VB.Menu mnuSwapSidebarAlignment 
         Caption         =   "Swap Sidebar Alignement"
         HelpContextID   =   53
      End
      Begin VB.Menu mnuSep13b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowNotes 
         Caption         =   "Show Notes"
         Checked         =   -1  'True
         HelpContextID   =   54
      End
   End
   Begin VB.Menu mnuCalcPopup 
      Caption         =   "CalcPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCalcCopy 
         Caption         =   "Copy"
         HelpContextID   =   55
      End
      Begin VB.Menu mnuCalcPaste 
         Caption         =   "Paste"
         HelpContextID   =   56
      End
      Begin VB.Menu mnuA 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalcClear 
         Caption         =   "Clear"
         HelpContextID   =   57
      End
   End
   Begin VB.Menu mnuTrayPopup 
      Caption         =   "TrayPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuRestore 
         Caption         =   "Open XSE"
         HelpContextID   =   58
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "Exit"
         HelpContextID   =   59
      End
   End
   Begin VB.Menu mnuEditPopup 
      Caption         =   "EditPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuSaveScript 
         Caption         =   "Save script"
         HelpContextID   =   13035
      End
      Begin VB.Menu mnuCompile 
         Caption         =   "Compile"
         HelpContextID   =   13036
      End
      Begin VB.Menu mnuDebug 
         Caption         =   "Debug"
         HelpContextID   =   13043
      End
      Begin VB.Menu mnuSep19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditUndo 
         Caption         =   "Undo"
         HelpContextID   =   10
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "Redo"
         HelpContextID   =   60
      End
      Begin VB.Menu mnuSep20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         HelpContextID   =   69
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         HelpContextID   =   70
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         HelpContextID   =   71
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
         HelpContextID   =   72
      End
      Begin VB.Menu mnuSep21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select All"
         HelpContextID   =   20
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Variables for Calc function
Private lOperand1 As Long
Private lOperand2 As Long
Private lOldOperand As Long
Private sOperator As String
Private sOldOperator As String
Private sMemory As String
Private CanOverwrite As Boolean
Private MultipleTimes As Boolean

Private IsPrevInstance As Boolean
Private CmdLineBusy As Boolean

Private ActualTabLimit As Integer
Private IsNewTab As Boolean
'Private m_TabsFocused As Boolean
Private m_Answer As Integer
Private m_JustLoaded As Boolean

Private CurrentFont As StdFont
Private lBackgroundColor As Long
Private lForegroundColor As Long
'Private sFont As String
'Private bFontBold As Byte
'Private bFontItalic As Byte
'Private bFontSize As Byte

Private Const IMAGE_ICON = 1
Private Const LR_LOADMAP3DCOLORS = &H1000&
Private Const LR_SHARED = &H8000&

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Dim TrayI As NOTIFYICONDATA

Private Const WM_MOUSEMOVE = &H200
'Private Const WM_LBUTTONDOWN = &H201
'Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONUP = &H205

Private Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function OpenIcon Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long

Private Const ERROR_ALREADY_EXISTS = 183&

Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As Any, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private lMinHeight As Long
Private lMinWidth As Long

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

Private Const WM_KILLFOCUS = &H8&
Private Const WM_GETMINMAXINFO = &H24&
Private Const WM_CONTEXTMENU = &H7B&

'Private Type CBT_CREATEWND
'    lpcs As Long
'    hWndInsertAfter As Long ' pointer to CreateStruct UDT
'End Type
'
'Private Type CREATESTRUCT
'    lpCreateParams As Long
'    hInstance As Long
'    hMenu As Long
'    hWndParent As Long
'    cy As Long
'    cx As Long
'    Y As Long
'    X As Long
'    Style As Long
'    lpszName As Long    ' pointer to window title
'    lpszClass As Long   ' atom or pointer to class name (always numeric)
'    ExStyle As Long
'End Type
'
'Private Const HCBT_CREATEWND As Long = 3
'Private Const SPI_GETWORKAREA As Long = 48
'Private Const WM_SHOWWINDOW As Long = &H18
'Private Const WM_WINDOWPOSCHANGING As Long = &H46
'
'Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Public WithEvents objTmr As APITimer
Attribute objTmr.VB_VarHelpID = -1
Private WithEvents cSubclasser As cSelfSubclasser
Attribute cSubclasser.VB_VarHelpID = -1

Public Property Get JustLoaded() As Boolean
    JustLoaded = m_JustLoaded
End Property

Public Property Let JustLoaded(bool As Boolean)
    m_JustLoaded = bool
End Property

Public Property Get Answer() As Integer
    Answer = m_Answer
End Property

Public Property Let Answer(iValue As Integer)
    m_Answer = iValue
End Property

'Public Property Get TabsFocused() As Boolean
'    TabsFocused = m_TabsFocused
'End Property
'
'Public Property Let TabsFocused(bool As Boolean)
'    m_TabsFocused = bool
'End Property

Public Sub SetupCueControl(ctl As Control, sCue As String)
  'assign the cue text to the control's edit
  'box, as well as to the tag property. Using
  'the tag property to store the cue prompt text
  'negates the requirement to maintain the text
  'in a separate array. If your application design
  'uses are using the tag property for another
  'purpose, such as to store the dirty text of
  'the control, then a string array must be maintained
  'along with a mechanism to identify the control
  'in order to assign the correct prompt to the
  'respective control.
   With ctl
      .ForeColor = vbButtonShadow
      
     'tag is set first to ensure
     'CheckCuePromptChange sets correct value
     'when the control's Change event fires
      .Tag = sCue
      .text = sCue
   End With
End Sub

Public Sub CheckCuePromptChange(ctl As Control)
   ctl.HelpContextID = Trim$(ctl.text) = ctl.Tag
End Sub

Public Sub CheckCuePromptOnFocus(ctl As Control)
   With ctl
      If .HelpContextID = True Then
         .text = vbNullString
         .ForeColor = vbWindowText
      Else
         .SelStart = 0
         .SelLength = Len(.text)
         .HelpContextID = False
      End If
   End With
End Sub

Public Sub CheckCuePromptBlur(ctl As Control)
   With ctl
      If LenB(Trim$(.text)) = 0 Then
         .text = .Tag
         .ForeColor = vbButtonShadow
         .HelpContextID = True
      End If
   End With
End Sub

Private Sub imgClose_Click()
    lblClose_Click
End Sub

Private Sub imgNext_Click()
    lblNext_Click
End Sub

Private Sub imgPrev_Click()
    lblPrev_Click
End Sub

Private Sub imgTabs_Click()
    lblTabs_Click
End Sub

Public Sub lblClose_Click()

    If Document(Tabs.SelectedTab).IsDirty Then
        If mnuIgnoreChanges.Checked = False Then
            Answer = MsgBox(LoadResString(1001), vbExclamation + vbYesNoCancel)
            Select Case Answer
                Case vbYes
                    If Document(Tabs.SelectedTab).Save = False Then
                        Exit Sub
                    End If
                Case vbCancel
                    Exit Sub
            End Select
        End If
    End If
    
    RemoveTab
    
End Sub

Private Sub lblPrev_Click()
    Tabs.SelectTab Tabs.SelectedTab - 1
End Sub

Private Sub lblNext_Click()
    Tabs.SelectTab Tabs.SelectedTab + 1
End Sub

Private Function SwapEndian(ByVal dw As Long) As Long
' by Mike D Sutton, Mike.Sutton@btclick.com, 20040914
  SwapEndian = _
      (((dw And &HFF000000) \ &H1000000) And &HFF&) Or _
      ((dw And &HFF0000) \ &H100&) Or _
      ((dw And &HFF00&) * &H100&) Or _
      ((dw And &H7F&) * &H1000000)
  If (dw And &H80&) Then SwapEndian = SwapEndian Or &H80000000
End Function

Private Function Ptr(ByVal sNumber As String) As String
    Ptr = Hex$(SwapEndian(CLng("&H" & sNumber) + &H8000000))
End Function

Private Function Ofs(ByVal sNumber As String) As String
    Ofs = Hex$(SwapEndian(CLng("&H" & sNumber)) - &H8000000)
End Function

Private Sub cmdOfs_Click()
    
    ErrorCheck
    
    If Val("&H" & Right$(txtDisplay.text, 2)) >= &H8 And Val("&H" & Right$(txtDisplay.text, 2)) <= &HF Then
        If Len(txtDisplay.text) >= 7 Then
            txtDisplay.text = Right$(String$(8, "0") & txtDisplay.text, 8)
            txtDisplay.text = Ofs(txtDisplay.text)
            CanOverwrite = True
        End If
    End If
    
    SidebarFocus
    
End Sub

Private Sub cmdPtr_Click()
    
    ErrorCheck
    
    If Len(txtDisplay.text) <= 7 Then
        txtDisplay.text = Right$(String$(6, "0") & txtDisplay.text, Len(txtDisplay.text))
        txtDisplay.text = Ptr(txtDisplay.text)
        CanOverwrite = True
    End If
    
    SidebarFocus
    
End Sub

Private Sub lblTabs_Click()
Dim i As Integer
    
    mnuTab(0).Checked = False
    
    For i = mnuTab.LBound + 1 To mnuTab.UBound
        mnuTab(i).Visible = False
        mnuTab(i).Checked = False
    Next i
    
    For i = 0 To Tabs.TabCount - 1
        mnuTab(i).Visible = True
        mnuTab(i).Caption = i + 1 & "." & vbTab & Tabs.TabText(i + 1)
    Next i
    
    mnuTab(Tabs.SelectedTab - 1).Checked = True
    PopupMenu mnuTabs, , , , mnuTab(Tabs.SelectedTab - 1)
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Dim MSG As Long
    
    MSG = x \ Screen.TwipsPerPixelX
    
    If MSG = WM_LBUTTONDBLCLK Then
        AppActivate Me.hWnd
    ElseIf MSG = WM_RBUTTONUP Then
        SetForegroundWindow Me.hWnd
        Me.PopupMenu mnuTrayPopup, , , , mnuRestore
    End If
    
End Sub

Private Sub GetTabLimit()

    ActualTabLimit = (Me.Width \ Screen.TwipsPerPixelX) / (lMinWidth \ 10)
        
    If ActualTabLimit > MaxTabLimit Then
        ActualTabLimit = MaxTabLimit
    ElseIf ActualTabLimit < 10 Then
        ActualTabLimit = 10
    End If

End Sub

Private Sub MDIForm_Resize()

    If IsPrevInstance Then Exit Sub
    
    If Me.WindowState <> vbMinimized Then
        GetTabLimit
    End If

    If mnuMinimizetoSystemTray.Checked Then
        If Me.WindowState = vbMinimized Then
            Me.Hide
            TrayI.cbSize = LenB(TrayI)
            TrayI.hWnd = Me.hWnd
            TrayI.uId = vbNull
            TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            TrayI.ucallbackMessage = WM_MOUSEMOVE
            TrayI.hIcon = LoadImage(App.hInstance, "AAA", IMAGE_ICON, 16, 16, LR_SHARED Or LR_LOADMAP3DCOLORS)
            TrayI.szTip = App.Title & vbNullChar
            Shell_NotifyIcon NIM_ADD, TrayI
        Else
            TrayI.cbSize = LenB(TrayI)
            TrayI.hWnd = Me.hWnd
            TrayI.uId = vbNull
            Shell_NotifyIcon NIM_DELETE, TrayI
        End If
    End If

End Sub

Private Sub mnuAbout_Click()
    Show2 frmAbout, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuAdvanceTrainer_Click()
    If GetExt(Document(Tabs.SelectedTab).LoadedFile) <> "gba" Then
        ShellExecute Me.hWnd, "open", App.Path & "\A-Trainer.exe", vbNullString, vbNullString, vbNormalFocus
    Else
        ShellExecute Me.hWnd, "open", App.Path & "\A-Trainer.exe", Document(Tabs.SelectedTab).LoadedFile, vbNullString, vbNormalFocus
    End If
End Sub

Private Sub mnuAssociate_Click()
    Show2 frmAssociate, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

'Private Sub mnuAutomaticallyCheck_Click()
    'mnuAutomaticallyCheck.Checked = Not mnuAutomaticallyCheck.Checked
'End Sub

Private Sub mnuAutoSave_Click()
    Show2 frmAutoSave, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuBackup_Click()
Dim iFileNum As Integer
Dim tmpFile As String
Dim sContents As String

    tmpFile = Document(Tabs.SelectedTab).LoadedFile
    Mid$(tmpFile, Len(tmpFile) - 2) = "bak"
    
    iFileNum = FreeFile

    Open Document(Tabs.SelectedTab).LoadedFile For Binary As #iFileNum
        sContents = SysAllocStringLen(vbNullString, LOF(iFileNum))
        Get #iFileNum, 1, sContents
    Close #iFileNum
    
    If FileExists(tmpFile) = False Then
    
        iFileNum = FreeFile

        DeleteFile (tmpFile)
        Open tmpFile For Binary As #iFileNum
            Put #iFileNum, , sContents
        Close #iFileNum

'        FileCopy Document(Tabs.selectedtab).loadedfile, tmpFile
    
    Else
        
        If MsgBox(LoadResString(1002), vbExclamation + vbYesNo) = vbYes Then

            iFileNum = FreeFile

            DeleteFile (tmpFile)
            Open tmpFile For Binary As #iFileNum
                Put #iFileNum, , sContents
            Close #iFileNum
'
'            FileCopy Document(Tabs.selectedtab).loadedfile, tmpFile
            
        End If
    
    End If
    
    sContents = vbNullString

End Sub

Private Sub mnuBatchCompiler_Click()
    Show2 frmBatch, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuBuiltinScripts_Click()
    Show2 frmTemplate, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuCalcClear_Click()
    cmdClear_Click
End Sub

Private Sub mnuCalcCopy_Click()
    SafeClipboardSet txtDisplay.text
End Sub

Private Sub mnuCalcPaste_Click()
    If optDecHex(0).Value Then
        If IsNumeric(Clipboard.GetText) Then
            txtDisplay.text = CLng(Clipboard.GetText)
        End If
    Else
        If IsHex(Clipboard.GetText) Then
            txtDisplay.text = Hex$(CLng("&H" & Right$(Clipboard.GetText, 8) Xor &HFFFFFFFF + 1))
        End If
    End If
End Sub

Private Sub mnuCheckNow_Click()
    Show2 frmUpdate, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuClearRecent_Click()
Dim i As Long

    RemoveIniSection App.Path & IniFile, "MRUList"
    RemoveIniSection App.Path & IniFile, "MRUPath"
    
    For i = 0 To mnuRecent.UBound
        mnuRecent(i).Caption = vbNullString
        mnuRecent(i).Visible = False
    Next i
        
    mnuRecentFiles.Enabled = False
    
End Sub

Private Sub mnuCommandHelp_Click()
    Show2 frmReference, frmMain, CBool(mnuAlwaysonTop.Checked)
    frmReference.cboList.ListIndex = 0
End Sub

Private Sub mnuCompile_Click()
    Document(Tabs.SelectedTab).Compile
End Sub

Private Sub mnuDebug_Click()
    IsDebugging = True
    Document(Tabs.SelectedTab).Compile
End Sub

Private Sub mnuDecompileOptions_Click()
    Show2 frmDecompileOptions, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuEdit_Click()

    If LenB(Document(Tabs.SelectedTab).txtCode.text) <> 0 Then
        If Document(Tabs.SelectedTab).txtCode.SelLength > 0 Then
            mnuCopy.Enabled = True
            mnuCut.Enabled = True
            mnuDelete.Enabled = True
        Else
            mnuCopy.Enabled = False
            mnuCut.Enabled = False
            mnuDelete.Enabled = False
        End If
        mnuFind.Enabled = True
    Else
        mnuCopy.Enabled = False
        mnuCut.Enabled = False
        mnuDelete.Enabled = False
        mnuFind.Enabled = False
    End If

    mnuPaste.Enabled = LenB(Clipboard.GetText) <> 0
    
End Sub

Private Sub mnuEditCopy_Click()
    mnuCopy_Click
End Sub

Private Sub mnuEditCut_Click()
    mnuCut_Click
End Sub

Private Sub mnuEditDelete_Click()
    mnuDelete_Click
End Sub

Private Sub mnuEditPaste_Click()
    mnuPaste_Click
End Sub

Private Sub mnuEditRedo_Click()
    mnuRedo_Click
End Sub

Private Sub mnuEditSelectAll_Click()
    mnuSelectAll_Click
End Sub

Private Sub mnuEditUndo_Click()
    mnuUndo_Click
End Sub

Private Sub mnuExpander_Click()
    Show2 frmExpander, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuFile_Click()

    mnuClose.Enabled = lblClose.Enabled
    
    If Tabs.TabCount >= ActualTabLimit Then
        mnuNew.Enabled = False
    Else
        mnuNew.Enabled = True
    End If
    
    If HasPrinters = False Then
        mnuPrint.Enabled = False
    End If
    
End Sub

Private Sub mnuFind_Click()
    Show2 frmFind, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuFormat_Click()
    If Document(Tabs.SelectedTab).txtCode.FontSize <= 8.25 Then
        mnuDecrease.Enabled = False
        mnuIncrease.Enabled = True
    ElseIf Document(Tabs.SelectedTab).txtCode.FontSize >= 72 Then
        mnuDecrease.Enabled = True
        mnuIncrease.Enabled = False
    End If
End Sub

Private Sub mnuFreeSpaceFinder_Click()
    If GetExt(Document(Tabs.SelectedTab).LoadedFile) <> "gba" Then
        ShellExecute Me.hWnd, "open", App.Path & "\FSF.exe", vbNullString, vbNullString, vbNormalFocus
    Else
        ShellExecute Me.hWnd, "open", App.Path & "\FSF.exe", Document(Tabs.SelectedTab).LoadedFile, vbNullString, vbNormalFocus
    End If
End Sub

Private Sub mnuGoto_Click()
    Show2 frmGoto, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuHelp_Click()
    mnuGuide.Enabled = FileExists(App.Path & "\Guide.chm")
End Sub

Private Sub mnuHexViewer_Click()
    Show2 frmHexViewer, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuIgnoreChanges_Click()
    mnuIgnoreChanges.Checked = Not mnuIgnoreChanges.Checked
End Sub

Private Sub mnuInlineCommandHelp_Click()
    mnuInlineCommandHelp.Checked = Not mnuInlineCommandHelp.Checked
End Sub

Public Sub mnuMenuBar_Click()
    
    mnuMenuBar.Checked = Not mnuMenuBar.Checked
    
    LockUpdate Me.hWnd
    
    mnuFile.Visible = CBool(mnuMenuBar.Checked)
    mnuEdit.Visible = CBool(mnuMenuBar.Checked)
    mnuFormat.Visible = CBool(mnuMenuBar.Checked)
    mnuOptions.Visible = CBool(mnuMenuBar.Checked)
    mnuTools.Visible = CBool(mnuMenuBar.Checked)
    mnuHelp.Visible = CBool(mnuMenuBar.Checked)
    
    MyDoEvents
    AppActivate Me.hWnd
    MyDoEvents
    UnlockUpdate Me.hWnd
    
End Sub

Private Sub mnuMinimizetoSystemTray_Click()
    mnuMinimizetoSystemTray.Checked = Not mnuMinimizetoSystemTray.Checked
End Sub

Private Sub mnuPrint_Click()
Dim oOpenDialog As clsCommonDialog

    Set oOpenDialog = New clsCommonDialog
    oOpenDialog.ShowPrinter Document(Tabs.SelectedTab), HidePrintToFile
    Set oOpenDialog = Nothing
    
End Sub

Private Sub mnuReadOnly_Click()
    
    mnuReadOnly.Checked = Not mnuReadOnly.Checked
    'SendMessage Document(Tabs.selectedtab).txtCode.hWnd, EM_SETREADONLY, CBool(mnuReadOnly.Checked), ByVal 0&
    Document(Tabs.SelectedTab).txtCode.Locked = CBool(mnuReadOnly.Checked)
    
    If mnuReadOnly.Checked Then
        mnuUndo.Enabled = False
        mnuRedo.Enabled = False
    Else
        mnuUndo.Enabled = Document(Tabs.SelectedTab).CanUndo
        mnuRedo.Enabled = Document(Tabs.SelectedTab).CanRedo
    End If
    
End Sub

Private Sub mnuRecent_Click(Index As Integer)
Dim sFile As String
Dim sPath As String
Dim sFilePath As String
    
    sFile = ReadIniString(App.Path & IniFile, "MRUList", (Index + 1))
    sPath = ReadIniString(App.Path & IniFile, "MRUPath", (Index + 1))
    sFilePath = sPath & sFile
    
    If LenB(sFilePath) <> 0 Then
        If FileExists(sFilePath) Then
            LoadNewDoc , sFilePath, True
        Else
            MsgBox LoadResString(1003), vbCritical
        End If
    Else
        MsgBox LoadResString(1003), vbCritical
    End If
    
End Sub

Private Sub mnuRedo_Click()
    Document(Tabs.SelectedTab).Redo
End Sub

Private Sub mnuRememberSize_Click()
    mnuRememberSize.Checked = Not mnuRememberSize.Checked
End Sub

Private Sub mnuRestore_Click()
    AppActivate Me.hWnd
End Sub

Private Sub mnuRevert_Click()
    If Document(Tabs.SelectedTab).IsDirty Then
        Document(Tabs.SelectedTab).LoadFile
        mnuRevert.Enabled = False
    End If
End Sub

Public Sub mnuSave_Click()

    Select Case GetExt(Document(Tabs.SelectedTab).LoadedFile)
    
        Case "rbc", "rbh", "rbt"
            Document(Tabs.SelectedTab).Save (True)
            Document(Tabs.SelectedTab).IsDirty = False
            StatusBar.PanelEnabled(4) = False

        Case Else
            mnuSaveAs_Click
            
    End Select
    
End Sub

Private Sub mnuSaveAs_Click()
    
    Document(Tabs.SelectedTab).Save
    
    If LenB(GetFileName(Document(Tabs.SelectedTab).LoadedFile)) <> 0 Then

        Select Case GetExt(Document(Tabs.SelectedTab).LoadedFile)
    
            Case "rbc", "rbh", "rbt"
            
                Document(Tabs.SelectedTab).Caption = GetFileName(Document(Tabs.SelectedTab).LoadedFile)
                Tabs.TabText(Tabs.SelectedTab) = GetFileName(Document(Tabs.SelectedTab).LoadedFile)
                Document(Tabs.SelectedTab).IsDirty = False
                StatusBar.PanelEnabled(4) = False
                
        End Select
        
    End If
            
End Sub

Private Sub mnuSaveScript_Click()
    mnuSave_Click
End Sub

Private Sub mnuTab_Click(Index As Integer)
    Tabs.SelectTab Index + 1
End Sub

Private Sub mnuTextAdjuster_Click()
    Show2 frmTextAdjuster, frmMain, CBool(frmMain.mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuTools_Click()
    mnuFreeSpaceFinder.Enabled = FileExists(App.Path & "\FSF.exe")
    mnuAdvanceTrainer.Enabled = FileExists(App.Path & "\A-Trainer.exe")
End Sub

Private Sub mnuTrayExit_Click()
    
    TrayI.cbSize = LenB(TrayI)
    TrayI.hWnd = Me.hWnd
    TrayI.uId = vbNull
    Shell_NotifyIcon NIM_DELETE, TrayI
    
    If mnuIgnoreChanges.Checked = False Then
        If NoChanges = False Then
            AppActivate Me.hWnd
        End If
    End If
    
    Unload Me
    
End Sub

Private Sub mnuUndo_Click()
    'SendMessage Document(Tabs.selectedtab).txtCode.hWnd, EM_UNDO, 0, ByVal 0&
    Document(Tabs.SelectedTab).Undo
End Sub

Private Sub SidebarFocus()
    
    On Error GoTo CantFocus
    picSidebar.SetFocus
    
CantFocus:
End Sub

Private Sub objTmr_Timer()
    
    Select Case GetExt(Document(Tabs.SelectedTab).LoadedFile)
    
        Case "rbc", "rbh", "rbt"
            
            If Document(Tabs.SelectedTab).IsDirty Then
                Document(Tabs.SelectedTab).Save (True)
                Document(Tabs.SelectedTab).IsDirty = False
                StatusBar.PanelEnabled(4) = False
            End If
            
    End Select
    
End Sub

Private Sub optDecHex_Click(Index As Integer)
Dim i As Integer

    Select Case Index
        
        Case 0
            
            For i = 10 To 15
                cmdNum(i).Enabled = False
            Next
            
            If Not IsHex(txtDisplay.text) Then
                cmdClear_Click
            End If
            
            txtDisplay.text = CLng("&H" & Right$(txtDisplay.text, 8) Xor &HFFFFFFFF + 1)
            cmdPtr.Enabled = False
            cmdOfs.Enabled = False
            cmdPlusMinus.Enabled = True
            optDecHex(0).Value = True
            
        Case 1
            
            For i = 10 To 15
                cmdNum(i).Enabled = True
            Next
            
            If Not IsNumeric(txtDisplay.text) Then
                cmdClear_Click
            End If
            
            txtDisplay.text = Hex$(txtDisplay.text)
            cmdPtr.Enabled = True
            cmdOfs.Enabled = True
            cmdPlusMinus.Enabled = False
            optDecHex(1).Value = True
            
    End Select
    
    CanOverwrite = True
    SidebarFocus
    
End Sub

Private Sub picSidebar_GotFocus()
    Document(Tabs.SelectedTab).HasFocus = False
End Sub

Public Sub picSidebar_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        
        Case vbKey6 ' Minus French
            If MapVirtualKey(191, &H0) = 52 Then
                If Shift = 0 Then
                    cmdSubtract_Click
                End If
            End If
            
        Case vbKey7 ' Divide Ita/Esp/Ger
            If MapVirtualKey(191, &H0) = 43 Then
                If Shift = vbShiftMask Then
                    cmdDivide_Click
                End If
            End If
            
        Case vbKey8 'Asterisk US
            If MapVirtualKey(191, &H0) = 53 Then
                If Shift = vbShiftMask Then
                    cmdMultiply_Click
                End If
            End If
            
        Case vbKeyAdd
            cmdAdd_Click
            
        Case vbKeySubtract
            cmdSubtract_Click
            
        Case vbKeyMultiply
            cmdMultiply_Click
            
        Case vbKeyDivide
            cmdDivide_Click
            
        Case 187
            
            Select Case MapVirtualKey(KeyCode, &H0)
            
                Case 27
                    ' Plus/Asterisk Ita/Esp/Ger
                    If Shift = 0 Then
                        cmdAdd_Click
                    ElseIf Shift = vbShiftMask Then
                        cmdMultiply_Click
                    End If
                    
                Case 13
                    ' Plus US
                    If Shift = vbShiftMask Then
                        cmdAdd_Click
                    End If
                
                Case 39
                    ' Plus Dutch
                    If Shift = 0 Then
                        cmdAdd_Click
                    End If
                
            End Select
            
        Case 189
            If MapVirtualKey(KeyCode, &H0) = 53 Then
                ' Minus Ita/Esp/Ger
                If Shift = 0 Then
                    cmdSubtract_Click
                End If
            ElseIf MapVirtualKey(KeyCode, &H0) = 12 Then
                ' Minus US
                If Shift = 0 Then
                    cmdSubtract_Click
                End If
            End If
            
        Case 186 ' Asterisk Dutch
            If MapVirtualKey(KeyCode, &H0) = 27 Then
                If Shift = 0 Then
                    cmdMultiply_Click
                End If
            End If
            
        Case 191
            If MapVirtualKey(KeyCode, &H0) = 53 Then
                ' Divide US
                cmdDivide_Click
            ElseIf MapVirtualKey(KeyCode, &H0) = 52 Then
                ' Divide French
                If Shift = vbShiftMask Then
                    cmdDivide_Click
                End If
            End If

        Case 219 ' Divide Dutch
            If MapVirtualKey(KeyCode, &H0) = 12 Then
                cmdDivide_Click
            End If
            
        Case 220 ' Asterisk French
            If MapVirtualKey(KeyCode, &H0) = 43 Then
                cmdMultiply_Click
            End If
            
        Case vbKeyDelete
            cmdClear_Click
            
    End Select
End Sub

Public Sub picSidebar_KeyPress(KeyCode As Integer)
    Select Case KeyCode
        
        Case vbKey0 To vbKey9
            cmdNum_Click KeyCode - vbKey0
            
        Case vbKeyA To vbKeyF
            If optDecHex(1).Value = True Then
                cmdNum_Click KeyCode - 55
            End If
            
        Case vbKeyC - 64 'Ctrl+C
            If GetKeyState(vbKeyControl) = -127 Then
                mnuCalcCopy_Click
            End If
            
        Case vbKeyV - 64 'Ctrl+V
            If GetKeyState(vbKeyControl) = -127 Then
                mnuCalcPaste_Click
            End If
            
        Case vbKeyReturn
            cmdEqual_Click
            
        Case vbKeyBack
            cmdCE_Click
            
    End Select
End Sub

Private Sub picSidebar_Resize()
    
    If Me.WindowState <> vbMinimized Then
        If picSidebar.ScaleHeight > 5 Then
            linSidebar.Y2 = picSidebar.ScaleHeight - 5
        End If
    End If
    
End Sub

Private Sub picStatusBar_Resize()
    
    If Me.WindowState <> vbMinimized Then
        StatusBar.Width = picStatusBar.ScaleWidth
        StatusBar.PanelWidth(1) = Int((StatusBar.Width \ 2) - 86) - 27
        If StatusBar.PanelWidth(1) > 121 Then
            StatusBar.PanelWidth(2) = StatusBar.PanelWidth(1) - 120
        End If
    End If

End Sub

Private Sub picTabs_Resize()

    If Me.WindowState <> vbMinimized Then
        If picTabs.ScaleWidth > 120 Then
            Tabs.Width = picTabs.ScaleWidth - 120
            picTabControl.Left = picTabs.ScaleWidth - picTabControl.Width - 140
        End If
    End If

End Sub

Private Sub ActivateChild(hWndChild As Long)
Const GW_CHILD = 5
'Const WM_MDIACTIVATE = &H222
Const WM_SETFOCUS = &H7
Const WM_MDIMAXIMIZE = &H225
'Const SW_SHOWMAXIMIZED = 3
'Dim wp As WINDOWPLACEMENT
    
    'GetWindowPlacement hWndChild, wp
    'wp.showCmd = SW_SHOWMAXIMIZED
    
    'SetWindowPlacement hWndChild, wp
    SendMessage GetWindow(Me.hWnd, GW_CHILD), WM_MDIMAXIMIZE, hWndChild, ByVal 0&
    
    On Error GoTo CantFocus
    Document(Tabs.SelectedTab).txtCode.SetFocus

CantFocus:
End Sub

Private Sub RestoreSize()
Dim iWindowState As Integer
Dim WindowHeight As Single
Dim WindowWidth As Single
Dim wp As WINDOWPLACEMENT

Const SW_HIDE = 0
'Const SW_SHOWNORMAL = 1
'Const SW_SHOWMAXIMIZED = 3

    wp.Length = Len(wp)
    mnuRememberSize.Checked = ReadIniString(App.Path & IniFile, "Options", "RememberSize", False)
    
    If mnuRememberSize.Checked = True Then
    
        iWindowState = ReadIniString(App.Path & IniFile, "Options", "WindowState", vbNormal)
        
        If iWindowState <> vbMaximized Then
            WindowHeight = ReadIniString(App.Path & IniFile, "Options", "WindowHeight", lMinHeight)
            WindowWidth = ReadIniString(App.Path & IniFile, "Options", "WindowWidth", lMinWidth)
            MoveWindow Me.hWnd, (Screen.Width - WindowWidth) \ (2 * Screen.TwipsPerPixelX), (Screen.Height - WindowHeight) \ (2 * Screen.TwipsPerPixelY), WindowWidth \ Screen.TwipsPerPixelX, WindowHeight \ Screen.TwipsPerPixelY, False
        End If
        
        GetWindowPlacement Me.hWnd, wp
        wp.showCmd = SW_HIDE
        
        SetWindowPlacement Me.hWnd, wp
        
        If iWindowState = vbMaximized Then
            Me.WindowState = vbMaximized
        End If
        
    End If
    
End Sub

Private Sub Tabs_TabClick(ByVal lTab As Long)
    
    If lTab > 1 Then
    
        If lTab < Tabs.TabCount Then
            lblNext.Enabled = True
            imgNext.Enabled = True
            imgNext.Picture = imgRightOn.Picture
        Else
            lblNext.Enabled = False
            imgNext.Enabled = False
            imgNext.Picture = imgRightOff.Picture
        End If
        
        lblPrev.Enabled = True
        imgPrev.Enabled = True
        imgPrev.Picture = imgLeftOn.Picture
    
    ElseIf lTab = Tabs.TabCount And Tabs.TabCount > 1 Then
    
        lblNext.Enabled = False
        lblPrev.Enabled = True
        imgNext.Enabled = False
        imgPrev.Enabled = True
        imgNext.Picture = imgRightOff.Picture
        imgPrev.Picture = imgLeftOn.Picture
    
    ElseIf lTab = 1 Then
    
        If Tabs.TabCount > 1 Then
            lblNext.Enabled = True
            imgNext.Enabled = True
            imgNext.Picture = imgRightOn.Picture
        Else
            lblNext.Enabled = False
            imgNext.Enabled = False
            imgNext.Picture = imgRightOff.Picture
        End If
        
        lblPrev.Enabled = False
        imgPrev.Enabled = False
        imgPrev.Picture = imgLeftOff.Picture
    
    Else
        lblNext.Enabled = False
        lblPrev.Enabled = False
        imgNext.Enabled = False
        imgPrev.Enabled = False
        imgNext.Picture = imgRightOff.Picture
        imgPrev.Picture = imgLeftOff.Picture
    End If
    
    If Tabs.TabCount > 1 Then
        lblClose.Enabled = True
        mnuClose.Enabled = True
        lblTabs.Enabled = True
        imgClose.Enabled = True
        imgTabs.Enabled = True
        imgClose.Picture = imgCrossOn.Picture
        imgTabs.Picture = imgDownOn.Picture
    Else
        lblClose.Enabled = False
        mnuClose.Enabled = False
        lblTabs.Enabled = False
        lblNext.Enabled = False
        lblPrev.Enabled = False
        imgClose.Enabled = False
        imgTabs.Enabled = False
        imgClose.Picture = imgCrossOff.Picture
        imgTabs.Picture = imgDownOff.Picture
    End If
    
    ActivateChild Document(lTab).hWnd
    
    If IsNewTab Then
        Document(lTab).txtCode.BackColor = lBackgroundColor
        Document(lTab).txtCode.ForeColor = lForegroundColor
        Document(lTab).txtCode.FontName = CurrentFont.name
        Document(lTab).txtCode.FontBold = CurrentFont.Bold
        Document(lTab).txtCode.FontItalic = CurrentFont.Italic
        Document(lTab).txtCode.FontSize = CurrentFont.SIZE
        IsNewTab = False
    End If

End Sub

Private Sub tmrKeyboard_Timer()
    GetKeyStatus
End Sub

Public Sub WelcomeText()
    
    On Error GoTo Hell
    
    If LenB(Username) <> 0 Then
        StatusBar.PanelCaption(1) = LoadResString(1004) & " " & Username & "! " & LoadResString(1005)
    Else
        StatusBar.PanelCaption(1) = LoadResString(1004) & "! " & LoadResString(1005)
    End If
    
Hell:
End Sub

Private Sub tmrRestore_Timer()
    WelcomeText
    tmrRestore.Enabled = False
End Sub

Private Sub txtCommandLine_Change()
Dim lFileEnd As Long
Dim sCommandLine As String
    
    If CmdLineBusy = True Then Exit Sub
    CmdLineBusy = True
    
    sCommandLine = txtCommandLine.text
    
    If InStrB(1, sCommandLine, ".", vbBinaryCompare) <> 0 Then
    
        If InStrB(sCommandLine, "\") = 0 Then
            sCommandLine = Replace(CurDir$ & "\" & sCommandLine, "\\", "\")
        End If
    
        DoReplace sCommandLine, """", vbNullString
        txtCommandLine.text = sCommandLine
        
    Else
        GoTo Finish
    End If

    If Not IsPrevInstance Then
    
        If LenB(sCommandLine) <> 0 Then
        
            lFileEnd = InStr(LCase$(sCommandLine), ".gba")
            
            If lFileEnd <> 0 Then
                
                lFileEnd = lFileEnd + 3
                Document(Tabs.SelectedTab).FileIndex = 1
                
                If JustLoaded Then
                    
                    Document(Tabs.SelectedTab).LoadedFile = Left$(sCommandLine, lFileEnd)
                    Document(Tabs.SelectedTab).cboFile.ListIndex = 1
                    
                    If Val("&H" & Mid$(sCommandLine, lFileEnd + 2)) = 0 Then
                        GoTo Finish
                    End If
                    
                Else
                
                    If LoadNewDoc(, Left$(sCommandLine, lFileEnd), False) = True Then
                        If Val("&H" & Mid$(sCommandLine, lFileEnd + 2)) = 0 Then
                            GoTo Finish
                        End If
                    Else
                        GoTo Finish
                    End If
                    
                End If
                
                Document(Tabs.SelectedTab).txtOffset.text = Mid$(sCommandLine, lFileEnd + 2)
                Decompile Left$(sCommandLine, lFileEnd), CLng("&H" & Mid$(sCommandLine, lFileEnd + 2))
                
            Else
                
                Select Case GetExt(sCommandLine)
                    
                    Case "rbc", "rbh", "rbt"
                    
                        Document(Tabs.SelectedTab).FileIndex = 0
                        
                        If JustLoaded Then
                            Document(Tabs.SelectedTab).LoadedFile = sCommandLine
                            Document(Tabs.SelectedTab).cboFile.ListIndex = 0
                            Document(Tabs.SelectedTab).Caption = GetFileName(sCommandLine)
                            Tabs.TabText(Tabs.SelectedTab) = GetFileName(sCommandLine)
                            Document(Tabs.SelectedTab).LoadFile
                        Else
                            Call LoadNewDoc(, sCommandLine, True)
                        End If
                    
                End Select
                
            End If
        End If
    End If
    
Finish:
    CmdLineBusy = False

End Sub

Public Function LoadNewDoc(Optional FirstTime As Boolean = False, Optional ByVal sFileName As String = vbNullString, Optional IsText As Boolean = True) As Boolean
Dim i As Integer
Dim sTemp As String

    If Tabs.TabCount = ActualTabLimit Then
        If LenB(sFileName) = 0 Then
            LoadNewDoc = False
            Exit Function
        Else
            If LenB(Document(Tabs.SelectedTab).txtCode.text) <> 0 Then
                LoadNewDoc = False
                Exit Function
            End If
        End If
    End If
    
    LoadNewDoc = True
    
    If LenB(sFileName) = 0 Then
        
        IsNewTab = True
        lDocCounter = lDocCounter + 1
        
        If FirstTime = False Then
        
            Set Document(Tabs.TabCount + 1) = New frmRubIDE
            Document(Tabs.TabCount + 1).Caption = CaptionBase & lDocCounter
            ShowLines Document(Tabs.TabCount + 1).txtCode, CBool(mnuLineNumbers.Checked)
            Load Document(Tabs.TabCount + 1)
            Document(Tabs.TabCount + 1).NewTabTemplate
            AddTab
            
            If LenB(txtCommandLine.text) <> 0 Then
                txtCommandLine.text = vbNullString
            End If
            
        Else
            
            Set Document(1) = New frmRubIDE
            Document(1).Caption = CaptionBase & lDocCounter
            LoadSettings
            ShowLines Document(1).txtCode, CBool(mnuLineNumbers.Checked)
            'SetActiveWindow Me.hWnd
            
            LockUpdate Me.hWnd
            MyDoEvents
            
            Load Document(1)
            Document(1).Show
            Document(1).WindowState = vbMaximized
            
            UnlockUpdate Me.hWnd
            
        End If
    
    Else
        
        If FileExists(sFileName) Then
            If FileLength(sFileName) = 0 Then
                MsgBox LoadResString(13030), vbExclamation
                Exit Function
            End If
        Else
            Exit Function
        End If
        
        If IsText Then
            
            sTemp = GetFileName(sFileName)
        
            For i = 1 To Tabs.TabCount
                
                If Tabs.TabText(i) = sTemp Then
                    
                    Tabs.SelectTab i
                    
                    If Document(i).LoadedFile = sFileName Then
                        Exit Function
                    End If
                    
                    Exit For
                    
                End If
                
            Next i
            
        End If
                                
        If LenB(Document(Tabs.SelectedTab).txtCode.text) = 0 Then

            If IsText Then
                Tabs.TabText(Tabs.SelectedTab) = sTemp
                Document(Tabs.SelectedTab).Caption = sTemp
            Else
                Document(Tabs.SelectedTab).Caption = Tabs.TabText(Tabs.SelectedTab)
            End If

        Else
            
            IsNewTab = True
            
            If Document(Tabs.TabCount + 1) Is Nothing Then
                Set Document(Tabs.TabCount + 1) = New frmRubIDE
            End If
                
            If IsText Then
                Document(Tabs.TabCount + 1).Caption = sTemp
                Document(Tabs.TabCount + 1).txtOffset.text = vbNullString
            Else
                lDocCounter = lDocCounter + 1
                Document(Tabs.TabCount + 1).Caption = CaptionBase & lDocCounter
            End If
            
            ShowLines Document(Tabs.TabCount + 1).txtCode, CBool(mnuLineNumbers.Checked)
            Load Document(Tabs.TabCount + 1)

            If IsText Then
                Call AddTab(sTemp)
            Else
                Call AddTab
            End If
                                
        End If
        
        If IsText Then
            Document(Tabs.SelectedTab).FileIndex = 0
            Document(Tabs.SelectedTab).LoadedFile = sFileName
            Document(Tabs.SelectedTab).cboFile.ListIndex = 0
            Document(Tabs.SelectedTab).LoadFile
        Else
            Document(Tabs.SelectedTab).FileIndex = 1
            Document(Tabs.SelectedTab).LoadedFile = sFileName
            Document(Tabs.SelectedTab).cboFile.ListIndex = 1
            MakeWritable sFileName
        End If
    
    End If
    
    If Tabs.TabCount > 1 Then
        picTabs.Height = 480
        picTabs.Visible = True
    End If
    
End Function


Private Sub cmdAnd_Click()
    
    ErrorCheck
    
    If optDecHex(0).Value Then
        lOperand1 = CLng(txtDisplay.text)
    Else
        lOperand1 = CLng("&H" & txtDisplay.text)
    End If
    
    sOperator = "and"
    CanOverwrite = True
    MultipleTimes = False
    SidebarFocus
    
End Sub

Private Sub cmdCE_Click()
    txtDisplay.text = 0
    SidebarFocus
End Sub

Private Sub cmdClear_Click()
    lOperand1 = 0
    lOperand2 = 0
    sOperator = vbNullString
    txtDisplay.text = 0
    MultipleTimes = False
    SidebarFocus
End Sub

Private Sub cmdDivide_Click()
    
    ErrorCheck
    
    Select Case sOperator
        Case "+", "-"
            lOldOperand = lOperand1
            sOldOperator = sOperator
        Case "*"
            MultipleTimes = False
            cmdEqual_Click
    End Select
    
    If optDecHex(0).Value Then
        lOperand1 = CLng(txtDisplay.text)
    Else
        lOperand1 = CLng("&H" & txtDisplay.text)
    End If
    
    sOperator = "/"
    CanOverwrite = True
    MultipleTimes = False
    SidebarFocus
    
End Sub

Private Function LeftShift(ByVal lNumber As Long, ByVal iNumBits As Integer) As Long
  LeftShift = lNumber * 2 ^ iNumBits
End Function

Private Function RightShift(ByVal lNumber As Long, ByVal iNumBits As Integer) As Long
  RightShift = lNumber \ 2 ^ iNumBits
End Function

Private Sub ErrorCheck()
    If Len(txtDisplay.text) > 11 Then
        txtDisplay.text = 0
    End If
End Sub


Private Sub cmdEqual_Click()
Dim lResult As Long

    On Error GoTo CalcError
    
    If Not MultipleTimes Then
        
        If optDecHex(0).Value Then
            lOperand2 = CLng(txtDisplay.text)
        Else
            lOperand2 = CLng("&H" & txtDisplay.text)
        End If
        
    End If
    
    MultipleTimes = True
    
    Select Case sOperator
        
        Case "+"
            lResult = lOperand1 + lOperand2
            lOperand1 = lResult
            
        Case "-"
            lResult = lOperand1 - lOperand2
            lOperand1 = lResult
            
        Case "*"
        
            lResult = lOperand1 * lOperand2
            lOperand1 = lResult
                            
            Select Case sOldOperator
                Case "+"
                    lOperand2 = lResult
                    lResult = lOldOperand + lResult
                    lOperand1 = lResult
                    sOperator = "+"
                Case "-"
                    lOperand2 = lResult
                    lResult = lOldOperand - lResult
                    lOperand1 = lResult
                    sOperator = "-"
            End Select
            
            sOldOperator = vbNullString

        Case "/"
        
            lResult = lOperand1 \ lOperand2
            lOperand1 = lResult
                            
            Select Case sOldOperator
                Case "+"
                    lOperand2 = lResult
                    lResult = lOldOperand + lResult
                    lOperand1 = lResult
                    sOperator = "+"
                Case "-"
                    lOperand2 = lResult
                    lResult = lOldOperand - lResult
                    lOperand1 = lResult
                    sOperator = "-"
            End Select
            
            sOldOperator = vbNullString
            
        Case "and"
            lResult = lOperand1 And lOperand2
            lOperand1 = lResult
            
        Case "or"
            lResult = lOperand1 Or lOperand2
            lOperand1 = lResult
            
        Case "xor"
            lResult = lOperand1 Xor lOperand2
            lOperand1 = lResult
            
        Case "ls"
            lResult = LeftShift(lOperand1, lOperand2)
            lOperand1 = lResult
            
        Case "rs"
            lResult = RightShift(lOperand1, lOperand2)
            lOperand1 = lResult
            
        Case Else
            Exit Sub
            
    End Select
    
    If optDecHex(0).Value Then
        txtDisplay.text = lResult
    Else
        txtDisplay.text = Hex$(lResult)
    End If
    
    CanOverwrite = True
    SidebarFocus
    Exit Sub

CalcError:

    txtDisplay.MaxLength = 40 'temporarily expand limit
    txtDisplay.text = LoadResString(1007) & Err.Description
    txtDisplay.MaxLength = 11 'restore limit
    lOperand1 = 0
    lOperand2 = 0
    lOldOperand = 0
    sOperator = vbNullString
    sOldOperator = vbNullString
    CanOverwrite = True
    MultipleTimes = False
    SidebarFocus

End Sub

Private Sub cmdLs_Click()
    
    ErrorCheck
    
    If optDecHex(0).Value Then
        lOperand1 = CLng(txtDisplay.text)
    Else
        lOperand1 = CLng("&H" & txtDisplay.text)
    End If
    
    sOperator = "ls"
    CanOverwrite = True
    MultipleTimes = False
    SidebarFocus
    
End Sub

Private Sub cmdSubtract_Click()
    
    ErrorCheck
    
    Select Case sOperator
        Case "*", "/"
            MultipleTimes = False
            cmdEqual_Click
    End Select
    
    If optDecHex(0).Value Then
        lOperand1 = CLng(txtDisplay.text)
    Else
        lOperand1 = CLng("&H" & txtDisplay.text)
    End If
    
    sOperator = "-"
    CanOverwrite = True
    MultipleTimes = False
    SidebarFocus
    
End Sub

Private Sub cmdMR_Click()
    txtDisplay.text = sMemory
    SidebarFocus
End Sub

Private Sub cmdMS_Click()
    ErrorCheck
    Select Case cmdMS.Caption
        Case "MS"
            If txtDisplay.text <> "0" Then
                txtMem.text = "M"
                sMemory = txtDisplay.text
                cmdMS.Caption = "MC"
                cmdMR.Enabled = True
            End If
        Case "MC"
            txtMem.text = vbNullString
            sMemory = vbNullString
            cmdMS.Caption = "MS"
            cmdMR.Enabled = False
    End Select
    SidebarFocus
End Sub

Private Sub cmdMultiply_Click()
    
    ErrorCheck
    
    Select Case sOperator
        Case "+", "-"
            lOldOperand = lOperand1
            sOldOperator = sOperator
        Case "/"
            MultipleTimes = False
            cmdEqual_Click
    End Select
    
    If optDecHex(0).Value Then
        lOperand1 = CLng(txtDisplay.text)
    Else
        lOperand1 = CLng("&H" & txtDisplay.text)
    End If
    
    sOperator = "*"
    CanOverwrite = True
    MultipleTimes = False
    SidebarFocus
    
End Sub

Private Sub cmdNot_Click()
    
    ErrorCheck
    
    'sOperator = "not"
    CanOverwrite = True
    
    If optDecHex(0).Value Then
        lOperand1 = CLng(txtDisplay.text)
        txtDisplay.text = Not lOperand1
    Else
        lOperand1 = CLng("&H" & txtDisplay.text)
        txtDisplay.text = Hex$(Not lOperand1)
    End If
    
    SidebarFocus
    
End Sub

Private Sub cmdNum_Click(Index As Integer)

    If txtDisplay.text = "0" Or CanOverwrite Then
        txtDisplay.text = cmdNum(Index).Caption
        If CanOverwrite Then CanOverwrite = False
    Else
        If optDecHex(0) Then
            If Len(txtDisplay.text) <= 8 Then
                txtDisplay.SelStart = Len(txtDisplay.text)
                txtDisplay.SelText = cmdNum(Index).Caption
            End If
        Else
            If Len(txtDisplay.text) <= 7 Then
                txtDisplay.SelStart = Len(txtDisplay.text)
                txtDisplay.SelText = cmdNum(Index).Caption
            End If
        End If
    End If
    
    SidebarFocus

End Sub

Private Sub cmdOr_Click()
    
    ErrorCheck
    
    If optDecHex(0).Value Then
        lOperand1 = CLng(txtDisplay.text)
    Else
        lOperand1 = CLng("&H" & txtDisplay.text)
    End If
    
    sOperator = "or"
    CanOverwrite = True
    MultipleTimes = False
    SidebarFocus
    
End Sub

Private Sub cmdAdd_Click()
    
    ErrorCheck
    
    Select Case sOperator
        Case "*", "/"
            MultipleTimes = False
            cmdEqual_Click
    End Select
    
    If optDecHex(0).Value Then
        lOperand1 = CLng(txtDisplay.text)
    Else
        lOperand1 = CLng("&H" & txtDisplay.text)
    End If
    
    sOperator = "+"
    CanOverwrite = True
    MultipleTimes = False
    SidebarFocus
    
End Sub

Private Sub cmdPlusMinus_Click()

    ErrorCheck
    
    If txtDisplay.text <> "0" Then
        If InStrB(1, txtDisplay.text, "-", vbBinaryCompare) = 0 Then
            txtDisplay.text = "-" & txtDisplay.text
        Else
            txtDisplay.text = Mid$(txtDisplay.text, 2)
        End If
    End If
    
    SidebarFocus
    
End Sub

Private Sub cmdRs_Click()
    
    ErrorCheck
    
    If optDecHex(0).Value Then
        lOperand1 = CLng(txtDisplay.text)
    Else
        lOperand1 = CLng("&H" & txtDisplay.text)
    End If
    
    sOperator = "rs"
    CanOverwrite = True
    MultipleTimes = False
    SidebarFocus
    
End Sub

Private Sub cmdXor_Click()
    
    ErrorCheck
    
    If optDecHex(0).Value Then
        lOperand1 = CLng(txtDisplay.text)
    Else
        lOperand1 = CLng("&H" & txtDisplay.text)
    End If
    
    sOperator = "xor"
    CanOverwrite = True
    MultipleTimes = False
    SidebarFocus
    
End Sub

Private Sub LoadSettings()
Dim i As Integer
Dim sIniPath As String

    sIniPath = App.Path & IniFile
    
    If ReadIniSection(sIniPath, "MRUList").Count > 0 Then
        For i = 1 To ReadIniSection(sIniPath, "MRUList").Count
            mnuRecent(i - 1).Caption = Replace(ReadIniString(sIniPath, "MRUList", i), "&", "&&")
            mnuRecent(i - 1).Visible = True
        Next i
        mnuRecentFiles.Enabled = True
    Else
        mnuRecentFiles.Enabled = False
    End If

    'Read settings from INI...
    mnuAlwaysonTop.Checked = CBool(ReadIniString(sIniPath, "Options", "AlwaysOnTop", 0))
    mnuMinimizetoSystemTray.Checked = CBool(ReadIniString(sIniPath, "Options", "MinimizeToSystemTray", 0))
    mnuIgnoreChanges.Checked = CBool(ReadIniString(sIniPath, "Options", "IgnoreChanges", 0))
    mnuLineNumbers.Checked = CBool(ReadIniString(sIniPath, "Options", "LineNumbers", 1))
    mnuMenuBar.Checked = CBool(ReadIniString(sIniPath, "Options", "MenuBar", 1))
    mnuStatusBar.Checked = CBool(ReadIniString(sIniPath, "Options", "StatusBar", 1))
    mnuSidebar.Checked = CBool(ReadIniString(sIniPath, "Options", "Sidebar", 1))
    mnuRight.Checked = CBool(ReadIniString(sIniPath, "Options", "SidebarAlignment", 1))
    mnuLeft.Checked = Not CBool(ReadIniString(sIniPath, "Options", "SidebarAlignment", 1))
    mnuShowNotes.Checked = CBool(ReadIniString(sIniPath, "Options", "ShowNotes", 1))
    mnuShowRecentFiles.Checked = CBool(ReadIniString(sIniPath, "Options", "ShowRecentFiles", 1))

    optDecHex(0).Value = CBool(ReadIniString(sIniPath, "Options", "CalcMode", 1))
    optDecHex(1).Value = Not optDecHex(0).Value

    mnuInlineCommandHelp.Checked = CBool(ReadIniString(sIniPath, "Options", "InlineHelp", 1))
    'mnuAutomaticallyCheck.Checked = CBool(0)

    NoLog = CBool(ReadIniString(sIniPath, "Options", "NoLog", 0))
    sEmulatorPath = ReadIniString(sIniPath, "Options", "EmulatorPath", vbNullString)

    iDecompileMode = ReadIniString(sIniPath, "Options", "DecompileMode", 1)
    iComments = ReadIniString(sIniPath, "Options", "Comments", 1)
    iRefactoring = ReadIniString(sIniPath, "Options", "Refactoring", 0)
    sCommentChar = ReadIniString(sIniPath, "Options", "CommentChar", "'")
    sRefactorDynamic = ReadIniString(sIniPath, "Options", "RefactorDynamic")

    If CBool(ReadIniString(sIniPath, "Options", "AutoSave", 0)) = 1 Then
        mnuAutoSave.Checked = True
        Me.objTmr.StartTimer 1000& * ReadIniLong(sIniPath, "Options", "SaveInterval", 60&)
    Else
        mnuAutoSave.Checked = False
    End If

    '...and apply them
    SetTopmostWindow Me.hWnd, CBool(mnuAlwaysonTop.Checked)

    mnuFile.Visible = CBool(mnuMenuBar.Checked)
    mnuEdit.Visible = CBool(mnuMenuBar.Checked)
    mnuFormat.Visible = CBool(mnuMenuBar.Checked)
    mnuOptions.Visible = CBool(mnuMenuBar.Checked)
    mnuTools.Visible = CBool(mnuMenuBar.Checked)
    mnuHelp.Visible = CBool(mnuMenuBar.Checked)

    picStatusBar.Visible = CBool(mnuStatusBar.Checked)
    picSidebar.Visible = CBool(mnuSidebar.Checked)
    picSidebar.Align = vbAlignLeft - CInt(mnuRight.Checked)

    If picSidebar.Align = vbAlignRight Then
        linSidebar.X1 = 2
        linSidebar.X2 = 2
    Else
        linSidebar.X1 = picSidebar.ScaleWidth - 2
        linSidebar.X2 = picSidebar.ScaleWidth - 2
    End If

    txtNotes.Visible = CBool(mnuShowNotes.Checked)
    mnuRecentFiles.Visible = CBool(mnuShowRecentFiles.Checked)

    lBackgroundColor = ReadIniString(sIniPath, "Format", "BackgroundColor", vbWindowBackground)
    lForegroundColor = ReadIniString(sIniPath, "Format", "ForegroundColor", vbWindowText)
    CurrentFont.name = ReadIniString(sIniPath, "Format", "FontName", "Courier New")
    CurrentFont.Bold = CBool(ReadIniString(sIniPath, "Format", "FontBold", 0))
    CurrentFont.Italic = CBool(ReadIniString(sIniPath, "Format", "FontItalic", 0))
    CurrentFont.SIZE = ReadIniString(sIniPath, "Format", "FontSize", 9)

    If lBackgroundColor = 0 And lForegroundColor = 0 Then
        mnuResetColors_Click
        If CurrentFont.name = "MS Sans Serif" And CurrentFont.SIZE = 8.25 Then
            mnuResetFont_Click
        End If
    Else
        Document(Tabs.SelectedTab).txtCode.BackColor = lBackgroundColor
        Document(Tabs.SelectedTab).txtCode.ForeColor = lForegroundColor
        Document(Tabs.SelectedTab).txtCode.FontName = CurrentFont.name
        Document(Tabs.SelectedTab).txtCode.FontBold = CurrentFont.Bold
        Document(Tabs.SelectedTab).txtCode.FontItalic = CurrentFont.Italic
        Document(Tabs.SelectedTab).txtCode.FontSize = CSng(CurrentFont.SIZE)
    End If
   
End Sub

Private Sub SaveSettings()
Dim sIniPath As String

    sIniPath = App.Path & IniFile
    
    WriteStringToIni sIniPath, "Options", "AlwaysOnTop", -CInt(mnuAlwaysonTop.Checked)
    WriteStringToIni sIniPath, "Options", "MinimizeToSystemTray", -CInt(mnuMinimizetoSystemTray.Checked)
    WriteStringToIni sIniPath, "Options", "RememberSize", -CInt(mnuRememberSize.Checked)
    
    If mnuRememberSize.Checked = True Then
    
        WriteStringToIni sIniPath, "Options", "WindowState", Me.WindowState
        
        If Me.WindowState <> vbMinimized Then
            WriteStringToIni sIniPath, "Options", "WindowHeight", Me.Height
            WriteStringToIni sIniPath, "Options", "WindowWidth", Me.Width
        End If
        
    Else
        WriteStringToIni sIniPath, "Options", "WindowState", vbNullString
        WriteStringToIni sIniPath, "Options", "WindowHeight", vbNullString
        WriteStringToIni sIniPath, "Options", "WindowWidth", vbNullString
    End If
    
    WriteStringToIni sIniPath, "Options", "IgnoreChanges", -CInt(mnuIgnoreChanges.Checked)
    WriteStringToIni sIniPath, "Options", "LineNumbers", -CInt(mnuLineNumbers.Checked)
    WriteStringToIni sIniPath, "Options", "MenuBar", -CInt(mnuMenuBar.Checked)
    WriteStringToIni sIniPath, "Options", "StatusBar", -CInt(mnuStatusBar.Checked)
    WriteStringToIni sIniPath, "Options", "Sidebar", -CInt(mnuSidebar.Checked)
    WriteStringToIni sIniPath, "Options", "SidebarAlignment", -CInt(mnuRight.Checked)
    WriteStringToIni sIniPath, "Options", "ShowNotes", -CInt(mnuShowNotes.Checked)
    WriteStringToIni sIniPath, "Options", "ShowRecentFiles", -CInt(mnuShowRecentFiles.Checked)
    WriteStringToIni sIniPath, "Options", "CalcMode", -CInt(optDecHex(0).Value)
    WriteStringToIni sIniPath, "Options", "InlineHelp", -CInt(mnuInlineCommandHelp.Checked)
    WriteStringToIni sIniPath, "Options", "AutoUpdateCheck", "0"
    WriteStringToIni sIniPath, "Options", "NoLog", -CInt(NoLog)
    WriteStringToIni sIniPath, "Options", "EmulatorPath", sEmulatorPath
    WriteStringToIni sIniPath, "Format", "BackgroundColor", lBackgroundColor
    WriteStringToIni sIniPath, "Format", "ForegroundColor", lForegroundColor
    WriteStringToIni sIniPath, "Format", "FontName", CurrentFont.name
    WriteStringToIni sIniPath, "Format", "FontBold", -CInt(CurrentFont.Bold)
    WriteStringToIni sIniPath, "Format", "FontItalic", -CInt(CurrentFont.Italic)
    WriteStringToIni sIniPath, "Format", "FontSize", CurrentFont.SIZE
    
End Sub

Private Sub LoadNotes()
    
    If FileExists(App.Path & "\Notes.txt") Then
        BlastText txtNotes, App.Path & "\Notes.txt"
    End If
    
    If LenB(txtNotes.text) = 0 Then
        SetupCueControl txtNotes, LoadResString(1008)
    End If
        
End Sub

Private Sub SaveNotes()
Dim iFileNum As Integer

    DeleteFile App.Path & "\Notes.txt"
    iFileNum = FreeFile
    
    If txtNotes.text <> LoadResString(1008) Then
        If LenB(txtNotes.text) <> 0 Then
            Open App.Path & "\Notes.txt" For Output As #iFileNum
                Print #iFileNum, txtNotes.text;
            Close #iFileNum
        End If
    End If
    
End Sub

Private Function PrevInstance() As Boolean
Dim lMutexHandle As Long

    lMutexHandle = CreateMutex(ByVal 0&, 1, App.Title & "_mutex")

    If Err.LastDllError = ERROR_ALREADY_EXISTS Then
        ' Release handles
        ReleaseMutex lMutexHandle
        CloseHandle lMutexHandle
        PrevInstance = True
    End If

End Function

Private Function GetHandleFromCaption(sCaption As String) As Long
Dim lhWndP As Long
Dim sStr As String
Const GW_HWNDNEXT = 2

    lhWndP = FindWindow(vbNullString, vbNullString) 'Parent window
    
    Do While lhWndP <> 0
        
        sStr = Space$(GetWindowTextLength(lhWndP) + 1)
        GetWindowText lhWndP, sStr, Len(sStr)
        
        If InStrB(1, sStr, sCaption, vbBinaryCompare) <> 0 Then
            GetHandleFromCaption = lhWndP
            Exit Do
        End If
        
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
        
    Loop
    
End Function

Private Sub AppActivate(hWndApp As Long)
Const SW_SHOW = 5
Const SW_RESTORE = 9

    If IsIconic(hWndApp) Then
        ShowWindow hWndApp, SW_RESTORE
    Else
        ShowWindow hWndApp, SW_SHOW
    End If
    
    If GetForegroundWindow() <> hWndApp Then
        OpenIcon hWndApp
        BringWindowToTop hWndApp
        SetForegroundWindow hWndApp
    End If
    
End Sub

Private Function IsEXE() As Boolean
    IsEXE = App.LogMode
End Function

Private Sub MDIForm_Load()
Dim sCaption As String
Dim lPrevHandle As Long
Dim i As Long
    
    If IsEXE Then
        IsPrevInstance = PrevInstance
    Else
        IsPrevInstance = App.PrevInstance
    End If
    
    ' See if there is a previous instance.
    If Not IsPrevInstance Then

        SetIcon Me.hWnd, "AAA"
        Localize Me
                
        Set cSubclasser = New cSelfSubclasser
        
        If cSubclasser.ssc_Subclass(Me.hWnd, 1, 1, Me) = True Then
            
            lMinHeight = Me.Height \ Screen.TwipsPerPixelX
            lMinWidth = Me.Width \ Screen.TwipsPerPixelY
            
            ' WM_GETMINMAXINFO used to prevent manual resizing; Before or After - doesn't matter
            cSubclasser.ssc_AddMsg Me.hWnd, eMsgWhen.MSG_AFTER, WM_GETMINMAXINFO

        End If
        
        If cSubclasser.ssc_Subclass(txtDisplay.hWnd, 2, 1, Me) = True Then
            cSubclasser.ssc_AddMsg txtDisplay.hWnd, eMsgWhen.MSG_BEFORE, eAllMessages.ALL_MESSAGES
        End If
        
        Set objTmr = New APITimer
        
        RestoreSize
        
        LoadNotes
        WelcomeText
        
        cmdMultiply.Caption = ChrW$(&HD7) ' ×
        cmdDivide.Caption = ChrW$(&HF7) ' ÷
        
        imgPrev.Picture = imgLeftOff.Picture
        imgNext.Picture = imgRightOff.Picture
        imgTabs.Picture = imgDownOff.Picture
        imgClose.Picture = imgCrossOff.Picture
        
        StatusBar.BackColor = picTabs.BackColor
        
        Set CurrentFont = New StdFont
        LoadNewDoc True
        
        If LenB(Command$) = 0 Then
            Document(1).NewTabTemplate
        End If
        
        InitCollections
        MaxCommand = &HE2
        LoadCommands
        
        JustLoaded = True
        txtCommandLine.text = Command$
        JustLoaded = False
        
        GetKeyStatus
        
        If cSubclasser.shk_SetHook(WH_KEYBOARD, , eMsgWhen.MSG_BEFORE, , 2, Me) = False Then
            tmrKeyboard.Enabled = True
        End If

        ReDim CustomColors(63) As Byte

        For i = LBound(CustomColors) To UBound(CustomColors)
            If (i + 1) Mod 4 <> 0 Then
                CustomColors(i) = 255
            End If
        Next i
        
        'If mnuAutomaticallyCheck.Checked Then
        '    Load frmUpdate
        '    frmUpdate.vcLiveUpdate.AutoCheck = True
        '    frmUpdate.vcLiveUpdate.NextStep
        'Else
        '    mnuCheckNow.Enabled = True
        'End If
      
    Else
        
        sCaption = Me.Caption
        Me.Caption = vbNullString

        ' See if we have command-line arguments.
        If LenB(Command$) <> 0 Then

'            ' Clear any existing DDE link.
'            txtCommandLine.LinkMode = vbLinkNone
'
'            ' Define link to source.
'            txtCommandLine.LinkTopic = App.Title & "|" & Me.LinkTopic
'
'            ' Establish manual link.
'            txtCommandLine.LinkMode = vbLinkManual
'
'            ' Push the command-line arguments
'            txtCommandLine.text = Command$
'            txtCommandLine.LinkPoke
            
            txtCommandLine.text = Command$
            WriteToPrevInstance txtCommandLine.text, sCaption
            MyDoEvents

        End If
        
        ' Activate the running instance
        lPrevHandle = GetHandleFromCaption(sCaption)
        
        If lPrevHandle > 0 Then
            AppActivate lPrevHandle
        End If

        Unload Me
        End
    
    End If
       
End Sub

Private Sub CloseAllForms()
    
    Do While Forms.Count > 1
        Unload Forms(Forms.Count - 1)
    Loop

End Sub

Private Function NoChanges() As Integer
Dim i As Long
Dim lCount As Long
Dim lSel As Long
    
    NoChanges = 1
    
    For i = 1 To Tabs.TabCount
        
        If Document(i).IsDirty Then
            
            NoChanges = 0
            lCount = lCount + 1&
            
            If lCount > 1& Then
                Exit Function
            Else
                lSel = i
            End If
            
        End If
        
    Next i
    
    If lCount = 1& Then
        NoChanges = lSel * 2&
    End If
    
End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim AskNoMore As Boolean
Dim iChanges As Integer
Dim i As Byte

    If IsPrevInstance Then Exit Sub
    
    If IsOpen("frmUpdate") Then
        If frmUpdate.vcLiveUpdate.IsUpdateReady Then
            Cancel = True
            Exit Sub
        End If
    End If
    
    iChanges = NoChanges
    
    If iChanges Mod 2 = 0 And mnuIgnoreChanges.Checked = False Then
        
        If Me.WindowState = vbMinimized Then
            AppActivate Me.hWnd
        End If
        
        Answer = -1
        
        If Tabs.TabCount > 1 Then
            
            If iChanges = 0 Then
            
                For i = 2 To Tabs.TabCount
                    
                    Tabs.SelectTab i
                    StatusBar.PanelEnabled(4) = Document(i).IsDirty
                    
                    If StatusBar.PanelEnabled(4) Then
                        
                        'Answer = MsgBox(LoadResString(1001), vbExclamation + vbYesNoCancel)
                        If AskNoMore = False Then
                            Show2 frmAsk, Me, CBool(mnuAlwaysonTop.Checked), vbModal
                        End If
                        
                        Select Case Answer
                            Case 0 'Cancel
                                Cancel = True
                                Exit Sub
                            Case 1 'Yes
                                If Document(Tabs.SelectedTab).Save = False Then
                                    Cancel = True
                                    Exit Sub
                                End If
                            Case 2 'Yes to All
                                If Document(Tabs.SelectedTab).Save = True Then
                                    AskNoMore = True
                                Else
                                    Cancel = True
                                    Exit Sub
                                End If
                            Case 4 'No to All
                                GoTo Continue
                        End Select
                        
                    End If
    
                    RemoveTab
                    
                Next i
                
                StatusBar.PanelEnabled(4) = Document(1).IsDirty
                
            Else
                
                Tabs.SelectTab iChanges \ 2
                StatusBar.PanelEnabled(4) = Document(iChanges \ 2).IsDirty
                
            End If
            
        End If

        If StatusBar.PanelEnabled(4) = True Then
            
            If AskNoMore = False Then
                
                Answer = MsgBox(LoadResString(1001), vbExclamation + vbYesNoCancel)
                
                Select Case Answer
                    Case vbCancel
                        Answer = 0
                    Case vbYes
                        Answer = 1
                    Case vbNo
                        Answer = 3
                End Select
                
            End If
            
            Select Case Answer
                Case 0 'Cancel
                    Cancel = True
                    Exit Sub
                Case 1 'Yes
                    If Document(Tabs.SelectedTab).Save = False Then
                        Cancel = True
                        Exit Sub
                    End If
                Case 2 'Yes to All
                    If Document(Tabs.SelectedTab).Save = True Then
                        AskNoMore = True
                    Else
                        Cancel = True
                        Exit Sub
                    End If
                Case 4 'No to All
                    GoTo Continue
            End Select
            
        End If
        
    End If
    
Continue:
    
    Me.Hide
    Tabs.RemoveAllTabs
    CloseAllForms
    
    SaveSettings
    SaveNotes
    
    Set objTmr = Nothing
    
    If tmrKeyboard.Enabled = False Then
        cSubclasser.shk_UnHook WH_KEYBOARD
    End If
    
    Set cSubclasser = Nothing
    FreeLibrary m_hMod
    
End Sub

Private Sub mnuAlwaysonTop_Click()
    mnuAlwaysonTop.Checked = Not mnuAlwaysonTop.Checked
    SetTopmostWindow Me.hWnd, CBool(mnuAlwaysonTop.Checked)
End Sub

Private Sub mnuBackgroundColor_Click()
Dim NewColor As Long
Dim i As Long
Dim oOpenDialog As clsCommonDialog

    Set oOpenDialog = New clsCommonDialog
    NewColor = oOpenDialog.ShowColor(Document(Tabs.SelectedTab).hWnd)
    Set oOpenDialog = Nothing
    
    If NewColor <> -1 Then
        For i = 1 To Tabs.TabCount
            Document(i).txtCode.BackColor = NewColor
        Next i
        lBackgroundColor = NewColor
    End If
    
    Redraw Document(Tabs.SelectedTab).hWnd
    
End Sub

Private Sub mnuClose_Click()
    If Tabs.TabCount > 1 Then
        Call lblClose_Click
    End If
End Sub

Private Sub mnuCopy_Click()
    SendMessage Document(Tabs.SelectedTab).txtCode.hWnd, WM_COPY, 0, ByVal 0&
End Sub

Public Sub mnuCut_Click()
    SendMessage Document(Tabs.SelectedTab).txtCode.hWnd, WM_CUT, 0, ByVal 0&
End Sub

Private Sub mnuDateTime_Click()
    
    SendMessageStr Document(Tabs.SelectedTab).txtCode.hWnd, EM_REPLACESEL, 1&, "'" & TimeValue(time$) & _
    ChrW$(32) & Mid$(Date$, 4, 2) & ChrW$(47) & Left$(Date$, 2) & ChrW$(47) & Right$(Date$, 4)
    
End Sub

Public Sub mnuDecrease_Click()
Dim i As Long
    
    LockUpdate Document(Tabs.SelectedTab).txtCode.hWnd

    For i = 1 To Tabs.TabCount
        Document(i).txtCode.FontSize = Document(i).txtCode.FontSize - 1
    Next i
    CurrentFont.SIZE = Document(Tabs.SelectedTab).txtCode.FontSize
    
    ShowLines Document(Tabs.SelectedTab).txtCode, False
    ShowLines Document(Tabs.SelectedTab).txtCode, CBool(mnuLineNumbers.Checked)
    UnlockUpdate Document(Tabs.SelectedTab).txtCode.hWnd

    mnuIncrease.Enabled = True
    mnuFormat_Click
   
End Sub

Private Sub mnuDelete_Click()
Const WM_CLEAR = &H303
    SendMessage Document(Tabs.SelectedTab).txtCode.hWnd, WM_CLEAR, 0, ByVal 0&
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFont_Click()
Dim NewFont As String
Dim i As Long
Dim bBold As Byte
Dim bItalic As Byte
Dim SIZE As Single
Dim oOpenDialog As clsCommonDialog
    
    Set oOpenDialog = New clsCommonDialog
    NewFont = oOpenDialog.ShowFont(Document(Tabs.SelectedTab))
    Set oOpenDialog = Nothing
    
    If LenB(NewFont) <> 0 Then
        LockUpdate Document(Tabs.SelectedTab).txtCode.hWnd
        bBold = CBool(Document(Tabs.SelectedTab).txtCode.FontBold)
        bItalic = CBool(Document(Tabs.SelectedTab).txtCode.FontItalic)
        SIZE = Document(Tabs.SelectedTab).txtCode.FontSize
        
        For i = 1 To Tabs.TabCount
            Document(i).txtCode.FontName = NewFont
            Document(i).txtCode.FontBold = bBold
            Document(i).txtCode.FontItalic = bItalic
            Document(i).txtCode.FontSize = SIZE
        Next i
        
        CurrentFont.name = NewFont
        CurrentFont.Bold = bBold
        CurrentFont.Italic = bItalic
        CurrentFont.SIZE = Document(Tabs.SelectedTab).txtCode.FontSize
        ShowLines Document(Tabs.SelectedTab).txtCode, False
        ShowLines Document(Tabs.SelectedTab).txtCode, CBool(mnuLineNumbers.Checked)
        UnlockUpdate Document(Tabs.SelectedTab).txtCode.hWnd
    End If
    
End Sub

Private Sub mnuForegroundColor_Click()
Dim NewColor As Long
Dim i As Long
Dim oOpenDialog As clsCommonDialog

    Set oOpenDialog = New clsCommonDialog
    NewColor = oOpenDialog.ShowColor(Document(Tabs.SelectedTab).hWnd)
    Set oOpenDialog = Nothing
    
    If NewColor <> -1 Then
        
        For i = 1 To Tabs.TabCount
            Document(i).txtCode.ForeColor = NewColor
        Next i
        
        lForegroundColor = NewColor
    End If
    
    Redraw Document(Tabs.SelectedTab).hWnd
    
End Sub

Private Sub mnuHideSidebar_Click()
    picSidebar.Visible = False
    mnuSidebar.Checked = False
End Sub

Public Sub mnuIncrease_Click()
Dim i As Long
    
    LockUpdate Document(Tabs.SelectedTab).txtCode.hWnd
    
    For i = 1 To Tabs.TabCount
        Document(i).txtCode.FontSize = Document(i).txtCode.FontSize + 1
    Next i
    CurrentFont.SIZE = Document(Tabs.SelectedTab).txtCode.FontSize
    
    ShowLines Document(Tabs.SelectedTab).txtCode, False
    ShowLines Document(Tabs.SelectedTab).txtCode, CBool(mnuLineNumbers.Checked)
    UnlockUpdate Document(Tabs.SelectedTab).txtCode.hWnd
    
    mnuDecrease.Enabled = True
    mnuFormat_Click
    
End Sub

Private Sub mnuLeft_Click()
    Select Case mnuLeft.Checked
        Case True
            mnuLeft.Checked = False
            mnuRight.Checked = True
            picSidebar.Align = vbAlignRight
            linSidebar.X1 = 2
            linSidebar.X2 = 2
        Case False
            mnuLeft.Checked = True
            mnuRight.Checked = False
            picSidebar.Align = vbAlignLeft
            linSidebar.X1 = picSidebar.ScaleWidth - 1
            linSidebar.X2 = picSidebar.ScaleWidth - 1
    End Select
End Sub

Private Sub mnuLineNumbers_Click()
    mnuLineNumbers.Checked = Not mnuLineNumbers.Checked
    ShowLines Document(Tabs.SelectedTab).txtCode, CBool(mnuLineNumbers.Checked)
End Sub

Private Sub mnuNew_Click()
    LoadNewDoc
End Sub

Private Sub mnuOpen_Click()
Dim sResult As String
Dim oOpenDialog As clsCommonDialog

    Set oOpenDialog = New clsCommonDialog
    sResult = oOpenDialog.ShowOpen(Me.hWnd, vbNullString, , "All Supported Files (*.rbc; *.rbh; *.rbt; *.gba)|*.rbc;*.rbh;*.rbt;*.gba|Script Files (*.rbc; *.rbh; *.rbt)|*.rbc;*.rbh;*.rbt|GameBoy Advance ROMs (*.gba)|*.gba|", FileMustExist Or PATHMUSTEXIST Or HideReadOnly)
    
    If LenB(sResult) <> 0 Then
    
        Debug.Print sResult & " " & GetExt(sResult)
        
        Select Case oOpenDialog.FilterIndex
        
            Case 1
                
                Select Case GetExt(sResult)
                
                    Case "rbc", "rbh", "rbt"
                        LoadNewDoc , sResult, True
                        
                    Case "gba"
                        LoadNewDoc , sResult, False
                    
                End Select
            
            Case 2
                LoadNewDoc , sResult, True
            
            Case 3
                LoadNewDoc , sResult, False
                      
        End Select
        
    End If
    
    Set oOpenDialog = Nothing
    
End Sub

'Private Sub mnuPageSetup_Click()
'    Dim oOpenDialog As clsCommonDialog
'    Set oOpenDialog = New clsCommonDialog
'    Call oOpenDialog.ShowPageSetupDlg(Document(Tabs.selectedtab).hWnd)
'End Sub

Private Sub mnuPaste_Click()
    SendMessage Document(Tabs.SelectedTab).txtCode.hWnd, WM_PASTE, 0, ByVal 0&
End Sub

'Private Sub mnuPrint_Click()
'    Dim oOpenDialog As clsCommonDialog
'    Set oOpenDialog = New clsCommonDialog
'    Call oOpenDialog.ShowPrinter(Document(Tabs.selectedtab))
'End Sub

Private Sub mnuGuide_Click()
    ShellExecute Me.hWnd, "open", App.Path & "\Guide.chm", vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub mnuResetColors_Click()
Dim i As Long

    For i = 1 To Tabs.TabCount
        Document(i).txtCode.BackColor = vbWindowBackground
        Document(i).txtCode.ForeColor = vbWindowText
    Next i
    
    lBackgroundColor = vbWindowBackground
    lForegroundColor = vbWindowText
    Redraw Document(Tabs.SelectedTab).hWnd
    
End Sub

Private Sub mnuResetFont_Click()
Dim i As Long
    
    LockUpdate Document(Tabs.SelectedTab).txtCode.hWnd

    For i = 1 To Tabs.TabCount
        Document(i).txtCode.FontName = "Courier New"
        Document(i).txtCode.FontSize = 9
        Document(i).txtCode.FontBold = False
        Document(i).txtCode.FontItalic = False
    Next i
    
    CurrentFont.name = "Courier New"
    CurrentFont.SIZE = 9
    CurrentFont.Bold = False
    CurrentFont.Italic = False
    
    ShowLines Document(Tabs.SelectedTab).txtCode, False
    ShowLines Document(Tabs.SelectedTab).txtCode, CBool(mnuLineNumbers.Checked)
    UnlockUpdate Document(Tabs.SelectedTab).txtCode.hWnd
    
    mnuIncrease.Enabled = True
    mnuDecrease.Enabled = True

End Sub

Private Sub mnuRight_Click()
    Select Case mnuRight.Checked
        Case True
            mnuRight.Checked = False
            mnuLeft.Checked = True
            picSidebar.Align = vbAlignLeft
            linSidebar.X1 = picSidebar.ScaleWidth - 2
            linSidebar.X2 = picSidebar.ScaleWidth - 2
        Case False
            mnuRight.Checked = True
            mnuLeft.Checked = False
            picSidebar.Align = vbAlignRight
            linSidebar.X1 = 2
            linSidebar.X2 = 2
    End Select
End Sub

Private Sub mnuSelectAll_Click()
    SendMessage Document(Tabs.SelectedTab).txtCode.hWnd, EM_SETSEL, 0&, ByVal -1
    SendMessage Document(Tabs.SelectedTab).txtCode.hWnd, EM_SCROLLCARET, 0&, ByVal 0
End Sub

Private Sub mnuShowNotes_Click()
    mnuShowNotes.Checked = Not mnuShowNotes.Checked
    txtNotes.Visible = CBool(mnuShowNotes.Checked)
End Sub

Private Sub mnuShowRecentFiles_Click()
    mnuShowRecentFiles.Checked = Not mnuShowRecentFiles.Checked
    mnuRecentFiles.Visible = CBool(mnuShowRecentFiles.Checked)
    mnuSep0.Visible = CBool(mnuShowRecentFiles.Checked)
End Sub

Private Sub mnuSidebar_Click()
    
    mnuSidebar.Checked = Not mnuSidebar.Checked
    picSidebar.Visible = CBool(mnuSidebar.Checked)
    mnuSidebarAlignment.Enabled = CBool(mnuSidebar.Checked)
    
End Sub

Private Sub mnuStatusBar_Click()
    mnuStatusBar.Checked = Not mnuStatusBar.Checked
    picStatusBar.Visible = CBool(mnuStatusBar.Checked)
End Sub

Private Sub mnuSwapSidebarAlignment_Click()
    Select Case picSidebar.Align
        Case vbAlignLeft
            mnuRight.Checked = True
            mnuLeft.Checked = False
            picSidebar.Align = vbAlignRight
            linSidebar.X1 = 2
            linSidebar.X2 = 2
        Case vbAlignRight
            mnuRight.Checked = False
            mnuLeft.Checked = True
            picSidebar.Align = vbAlignLeft
            linSidebar.X1 = picSidebar.ScaleWidth - 1
            linSidebar.X2 = picSidebar.ScaleWidth - 1
    End Select
End Sub

Private Sub picSidebar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If Button = vbRightButton Then
        PopupMenu mnuSidebarPopup
    End If
    
    SidebarFocus
    
End Sub

Private Sub txtDisplay_GotFocus()
    SendMessage txtDisplay.hWnd, WM_KILLFOCUS, 0&, ByVal 0&
    SidebarFocus
End Sub

Private Sub txtDisplay_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    If Button = vbLeftButton Then
        SidebarFocus
    End If

End Sub

Private Sub StatusBar_DblClick(lPanelNumber As Long)
    Select Case lPanelNumber
        Case 5
            StatusBar.PanelEnabled(lPanelNumber) = Not StatusBar.PanelEnabled(lPanelNumber)
            ToggleKey vbKeyCapital, StatusBar.PanelEnabled(lPanelNumber)
        Case 6
            StatusBar.PanelEnabled(lPanelNumber) = Not StatusBar.PanelEnabled(lPanelNumber)
            ToggleKey vbKeyNumlock, StatusBar.PanelEnabled(lPanelNumber)
        Case 7
            StatusBar.PanelEnabled(lPanelNumber) = Not StatusBar.PanelEnabled(lPanelNumber)
            ToggleKey vbKeyScrollLock, StatusBar.PanelEnabled(lPanelNumber)
    End Select
End Sub

Private Sub txtNotes_Change()
    CheckCuePromptChange txtNotes
End Sub

Private Sub txtNotes_GotFocus()
    CheckCuePromptOnFocus txtNotes
End Sub

Private Sub txtNotes_LostFocus()

    CheckCuePromptBlur txtNotes
    
    If LenB(txtNotes.text) = 0 Then
        SetupCueControl txtNotes, LoadResString(1008)
    End If
    
End Sub

Private Sub txtNotes_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
    
    If Shift = vbCtrlMask Then
        ReleaseCapture
        SendMessage txtNotes.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
    End If
    
End Sub

Private Sub AddTab(Optional sCaption As String = vbNullString)

    If Tabs.TabCount < ActualTabLimit Then
        
        If LenB(sCaption) = 0 Then
            Tabs.AddTab CaptionBase & lDocCounter
        Else
            Tabs.AddTab sCaption
        End If
        
    End If

End Sub

Private Sub FixIndexes(ByVal lIndex As Long)
Dim i As Long
    
    If lIndex + 1 <= Tabs.TabCount Then
        
        For i = lIndex + 1 To UBound(Document)
            
            RtlMoveMemory ByVal VarPtr(Document(0)) + ((i - 1) * 4), ByVal VarPtr(Document(0)) + (i * 4), 4
            RtlMoveMemory ByVal VarPtr(Document(0)) + (i * 4), 0&, 4
            
            If Document(i + 1) Is Nothing Then
                Exit For
            End If
            
        Next i
    
    End If
    
End Sub

Private Sub RemoveTab()
    
    Unload Document(Tabs.SelectedTab)
    FixIndexes Tabs.SelectedTab
    
    Tabs.RemoveTab Tabs.SelectedTab
    
    If Tabs.TabCount = 1 Then
        picTabs.Height = 0
        picTabs.Visible = False
    End If
    
    If LenB(txtCommandLine.text) <> 0 Then
        txtCommandLine.text = vbNullString
    End If
    
End Sub

'Private Function PositionDialog(dlgWidth As Long, dlgHeight As Long, dlgLeft As Long, dlgTop As Long) As Boolean
'
'    ' Example of centering dialog boxes
'    ' The myWndProc & myHookProc, near end of this module, calls this procedure,
'    ' passing the dialog window's width, height, left & top.
'    ' We simply modify the left,top coords
'
'    ' Remember that APIs use pixels, therefore we need to convert
'    ' VB's vbTwips to vbPixels in order to provide accurate coords.
'    If dlgWidth > 0 And dlgHeight > 0 Then
'        ' when centering, check for width & height because they could be zero, believe it or not
'
'        Dim wRect As RECT
'        SystemParametersInfo SPI_GETWORKAREA, 0&, wRect, 0&
'
'        dlgLeft = ((Me.Width \ Screen.TwipsPerPixelX) - dlgWidth) \ 2 + Me.Left \ Screen.TwipsPerPixelX
'        ' simple check to prevent dialog from displaying off the screen (horizontally)
'        If (dlgLeft + dlgWidth) > wRect.Right Then dlgLeft = wRect.Right - dlgWidth
'        If dlgLeft < wRect.Left Then dlgLeft = wRect.Left
'
'        dlgTop = ((Me.Height \ Screen.TwipsPerPixelY) - dlgHeight) \ 2 + Me.Top \ Screen.TwipsPerPixelY
'        ' simple check to prevent dialog from displaying off the screen (vertically)
'        If (dlgTop + dlgHeight) > wRect.Bottom Then dlgTop = wRect.Bottom - dlgHeight
'        If dlgTop < wRect.Top Then dlgTop = wRect.Top
'
'        PositionDialog = True
'
'    End If
'
'End Function
'
' ordinal #2 'Hook procedure used to center a dialog window
'Private Sub myHookProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lHookType As eHookType, ByRef lParamUser As Long)
'*************************************************************************************************
' http://msdn2.microsoft.com/en-us/library/ms644990.aspx
'* bBefore    - Indicates whether the callback is before or after the next hook in chain.
'* bHandled   - In a before next hook in chain callback, setting bHandled to True will prevent the
'*              message being passed to the next hook in chain and (if set to do so).
'* lReturn    - Return value. For Before messages, set per the MSDN documentation for the hook type
'* nCode      - A code the hook procedure uses to determine how to process the message
'* wParam     - Message related data, hook type specific
'* lParam     - Message related data, hook type specific
'* lHookType  - Type of hook calling this callback
'* lParamUser - User-defined callback parameter. Change vartype as needed (i.e., Object, UDT, etc)
'*************************************************************************************************
'
''http://msdn2.microsoft.com/en-us/library/ms644977.aspx
'    If lHookType = WH_CBT Then
'        If nCode = HCBT_CREATEWND Then  ' flag indicating window is being created
'
'            Dim wcw As CREATESTRUCT
'            Dim hcw As CBT_CREATEWND
'            ' get the hcbt_createwnd structure
'            RtlMoveMemory hcw, ByVal lParam, Len(hcw)
'
'            If hcw.lpcs Then    ' pointer to a createstruct
'                RtlMoveMemory wcw, ByVal hcw.lpcs, Len(wcw)   ' get that structure
'
'                If wcw.lpszClass = 32770 Then ' dialog class name atom
'                    ' not all dialogs are created equal :)
'                    ' messageboxes can be positioned here, while other dialogs
'                    ' can change position coords when this is received
'                    ' and when it is finally shown....
'
'                    ' by trying to adjust size here and also subclassing
'                    ' the window, we can catch it either way
'
'                    ' call local procedure to position dialog & save results
'                    If PositionDialog(wcw.cx, wcw.cy, wcw.X, wcw.Y) Then
'                        RtlMoveMemory ByVal hcw.lpcs, wcw, Len(wcw)
'                    End If
'
'                    ' start subclassing the dialog window. wParam is the handle
'                    ' we want the subclass procedure at ordinal #1 in our form (myWndProc)
'                    ' We also include the optional parameter to let the window procedure know messages are for this example
'                    If cSubclasser.ssc_Subclass(wParam, ByVal 3, 1, Me) Then
'                        cSubclasser.ssc_AddMsg wParam, eMsgWhen.MSG_AFTER, WM_SHOWWINDOW, WM_WINDOWPOSCHANGING
'                    End If
'
'                    ' we can unhook now
'                    cSubclasser.shk_UnHook lHookType
'                End If
'
'            End If
'
'        End If
'    End If
'End Sub

'- ordinal #2
Private Sub myHookProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal lHookType As eHookType, ByRef lParamUser As Long)
    
    If nCode = 0 Then ' HC_ACTION
    
        If lHookType = eHookType.WH_KEYBOARD Then
        
            If ((lParam And &HC0000000) <> 0) Then
            
                Select Case wParam
                    
                    Case vbKeyCapital, vbKeyNumlock, vbKeyScrollLock
                        
                        GetKeyStatus
                        
                    Case vbKeyF6, vbKeyTab
                        
                        If GetKeyState(vbKeyControl) < -1 Then
                            
                            If GetKeyState(vbKeyShift) >= 0 Then
                                
                                If Tabs.TabCount > 1 Then
                                    
                                    If Tabs.SelectedTab < Tabs.TabCount Then
                                        Tabs.SelectTab Tabs.SelectedTab + 1
                                    Else
                                        Tabs.SelectTab 1
                                    End If
                                    
                                End If
                                
                            Else
                                
                               If Tabs.TabCount > 1 Then
                                    
                                    If Tabs.SelectedTab > 1 Then
                                        Tabs.SelectTab Tabs.SelectedTab - 1
                                    Else
                                       Tabs.SelectTab Tabs.TabCount
                                    End If
    
                                End If
                                
                            End If
                            
                            lReturn = 1
                            
                        End If
                        
                End Select
            End If
        End If
    End If
    
End Sub

'- ordinal #1
Private Sub myWndProc(ByVal bBefore As Boolean, ByRef bHandled As Boolean, ByRef lReturn As Long, ByVal lng_hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef lParamUser As Long)
'*************************************************************************************************
'* bBefore    - Indicates whether the callback is before or after the original WndProc. Usually
'*              you will know unless the callback for the uMsg value is specified as
'*              MSG_BEFORE_AFTER (both before and after the original WndProc).
'* bHandled   - In a before original WndProc callback, setting bHandled to True will prevent the
'*              message being passed to the original WndProc.
'*              Not applicable with After messages
'* lReturn    - WndProc return value. Set as per the MSDN documentation for the message value,
'*              and/or, in an after the original WndProc callback, act on the return value as set
'*              by the original WndProc.
'* lng_hWnd   - Window handle.
'* uMsg       - Message value.
'* wParam     - Message related data.
'* lParam     - Message related data.
'* lParamUser - User-defined callback parameter
'*************************************************************************************************
        
        Select Case lParamUser
        
            Case 1
        
                If uMsg = WM_GETMINMAXINFO Then
                    
                    Dim mmi As MINMAXINFO
                    
                    RtlMoveMemory mmi, ByVal lParam, Len(mmi)  ' get suggested min/max data
                    
                    mmi.ptMinTrackSize.x = lMinWidth        ' set our min, unmaximized, size
                    mmi.ptMinTrackSize.Y = lMinHeight
                    
                    RtlMoveMemory ByVal lParam, mmi, Len(mmi)  ' override
                
                    bHandled = True
                    lReturn = 1
                    
                End If
                
            Case 2
            
                If uMsg = WM_CONTEXTMENU Then
                
                    PopupMenu mnuCalcPopup
                    
                    bHandled = True
                    lReturn = 1
                    
                End If
            
        End Select

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
' *************************************************************
        
End Sub
