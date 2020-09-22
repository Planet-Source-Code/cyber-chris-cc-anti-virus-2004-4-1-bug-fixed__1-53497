VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "CCAntivirus 2004"
   ClientHeight    =   3735
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSearch.frx":0000
   ScaleHeight     =   3735
   ScaleWidth      =   8640
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   8040
      Top             =   120
   End
   Begin CCAntivir2004.Hyperlink lblBug 
      Height          =   255
      Left            =   5280
      TabIndex        =   35
      Top             =   3480
      Width           =   3015
      _ExtentX        =   3201
      _ExtentY        =   450
      ForeColorIdle   =   16711680
      ForeColorMouse  =   255
      BackColor       =   16777215
      Caption         =   " Report a bug"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.PictureBox picScan 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   2895
      TabIndex        =   18
      Top             =   120
      Width           =   2895
      Begin VB.PictureBox picFastSearchx 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   1
         Left            =   120
         Picture         =   "frmSearch.frx":8A12
         ScaleHeight     =   615
         ScaleWidth      =   2460
         TabIndex        =   21
         Top             =   1560
         Width           =   2460
         Begin VB.Label lblffs 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fast file search"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   495
            TabIndex        =   42
            Top             =   180
            Width           =   1770
         End
      End
      Begin VB.PictureBox picPathsearch 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":C300
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   20
         Top             =   960
         Width           =   2055
         Begin VB.Label lblCFP 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Full path search"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   690
            TabIndex        =   41
            Top             =   180
            Width           =   1320
         End
      End
      Begin VB.PictureBox picFileSearch 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":FE0A
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   19
         Top             =   360
         Width           =   2055
         Begin VB.Label lblSif 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Search in files"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   525
            TabIndex        =   40
            Top             =   195
            Width           =   1485
         End
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   2
         X1              =   2880
         X2              =   2880
         Y1              =   240
         Y2              =   2280
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   1
         X1              =   2880
         X2              =   0
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   0
         X1              =   0
         X2              =   0
         Y1              =   240
         Y2              =   2280
      End
      Begin VB.Label lblFileScan 
         Caption         =   "   File Scan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox picOther 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   270
      Left            =   120
      ScaleHeight     =   270
      ScaleWidth      =   2895
      TabIndex        =   23
      Top             =   465
      Width           =   2895
      Begin VB.PictureBox picUpdate 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":1395C
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   25
         Top             =   960
         Width           =   2055
         Begin VB.Label lblupdate 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Update"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   525
            TabIndex        =   38
            Top             =   165
            Width           =   1485
         End
      End
      Begin VB.PictureBox picSec 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":15AFA
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   24
         Top             =   360
         Width           =   2055
         Begin VB.Label lblSecured 
            BackColor       =   &H00FFFFFF&
            Caption         =   "About"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   450
            TabIndex        =   39
            Top             =   165
            Width           =   1365
         End
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   5
         X1              =   2880
         X2              =   2880
         Y1              =   120
         Y2              =   1680
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   4
         X1              =   2880
         X2              =   0
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   3
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1680
      End
      Begin VB.Label lblOther 
         Caption         =   "   Extra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.PictureBox picHelpAbout 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   2895
      TabIndex        =   27
      Top             =   840
      Width           =   2895
      Begin VB.PictureBox picAbout 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":18E6C
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   30
         Top             =   960
         Width           =   2055
         Begin VB.Label lblAbout 
            BackColor       =   &H00FFFFFF&
            Caption         =   "About"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.PictureBox picHelp 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         Picture         =   "frmSearch.frx":1C4FE
         ScaleHeight     =   615
         ScaleWidth      =   2055
         TabIndex        =   29
         Top             =   360
         Width           =   2055
         Begin VB.Label lblHelp 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Help"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   510
            TabIndex        =   37
            Top             =   165
            Width           =   1095
         End
      End
      Begin VB.Label lblAbuthelp 
         Caption         =   "   Help / About"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   2895
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   8
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   1680
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   7
         X1              =   2850
         X2              =   -30
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line lineScanFiles 
         BorderColor     =   &H8000000F&
         Index           =   6
         X1              =   2880
         X2              =   2880
         Y1              =   120
         Y2              =   1680
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   11400
      ScaleHeight     =   2955
      ScaleWidth      =   2655
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   2715
      Begin VB.CommandButton cmdScan 
         Caption         =   "Scan the selected File"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   45
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Checksum:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filesize:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   705
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filename:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   60
         Width           =   825
      End
      Begin VB.Label lblFileName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " "
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   300
         Width           =   45
      End
   End
   Begin VB.Label lblLogFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Log"
      Height          =   255
      Left            =   7560
      TabIndex        =   45
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   1
      X1              =   8640
      X2              =   3120
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblText 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      Caption         =   "Active File Monitoring"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   3240
      TabIndex        =   44
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   6120
      TabIndex        =   43
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   6120
      TabIndex        =   34
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblText 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      Caption         =   "Logging:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   3240
      TabIndex        =   33
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   6120
      TabIndex        =   32
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblText 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      Caption         =   "Files in quarintine:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   3240
      TabIndex        =   31
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   8
      X1              =   6840
      X2              =   6840
      Y1              =   -720
      Y2              =   0
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   7
      X1              =   6360
      X2              =   6360
      Y1              =   -720
      Y2              =   0
   End
   Begin VB.Label lblText 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      Caption         =   "Run on startup:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3240
      TabIndex        =   9
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   6120
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   6120
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblText 
      Alignment       =   1  'Rechts
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tray window:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5520
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Signatures:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   3720
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   0
      X1              =   3120
      X2              =   3120
      Y1              =   4080
      Y2              =   -120
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   2
      X1              =   8640
      X2              =   3120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anti Virus Definitions:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3720
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "0.0.0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5520
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   1
      Top             =   3240
      Width           =   3255
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Files checked:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuLMainWindow 
         Caption         =   "Load Main Window"
      End
      Begin VB.Menu mnuSF 
         Caption         =   "Scan File"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit CC Antivir"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Private WithEvents X     As cCommonDialog
Attribute X.VB_VarHelpID = -1
Private CurrentFile      As String
Private sPicScan         As pStatus
Private SpicOther        As pStatus
Private sHelpAbout       As pStatus
Dim isoverlabel  As Boolean
Dim iewindow As InternetExplorer
Private currentwindows     As New ShellWindows


Private Sub cmdScan_Click()

    CheckFile (CurrentFile)

End Sub

Private Sub Form_Load()
On Error GoTo err:
    Call lblFileScan_Click
    TranslateLabel lblBug, 115
    TranslateLabel lblText(6), 116
    TranslateLabel lblText(9), 117
    TranslateLabel lblText(17), 118
    TranslateLabel lblText(15), 119
    TranslateLabel lblText(4), 120
    TranslateLabel lblText(2), 121
    TranslateLabel lblText(0), 123
    TranslateLabel lblFileScan, 125
    TranslateLabel lblOther, 126
    TranslateLabel lblAbuthelp, 127
    TranslateLabel lblAbout, 130
    TranslateLabel lblText(20), 152
    TranslateLabel lblupdate, 143
    TranslateLabel lblSecured, 145
    TranslateLabel lblSif, 146
    TranslateLabel lblCFP, 147
    TranslateLabel lblffs, 148
    
    logType(1) = "Virus found"
    logType(2) = "Virus action"
    logType(3) = "Error"
    Set X = New cCommonDialog
    Set ccClass = X
    frmMain.Cls
    BuildUI
    sPicScan = Max
    SpicOther = Min
    sHelpAbout = Min
    If DateDiff("d", lblText(3).Caption, CDate(date)) > 5 Then
        If DateDiff("d", GetSetting(AV.AVname, "Settings", "RemindLater", date), date) >= 0 Then
            frmAutoUpdate.Show , Me
        End If
    End If
    Debug.Print DateDiff("d", GetSetting(AV.AVname, "Settings", "RemindLater", date), date)

Exit Sub
err:
 ErrorFunc err.Number, err.Description, "frMmain.Startup"
End Sub





Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isoverlabel = True
End Sub

Private Sub Form_Unload(cancel As Integer)
    
    End

End Sub

Private Sub lblAbout_Click()

    picAbout_Click

End Sub

Private Sub lblAbuthelp_Click()

  Dim temp As Long

    If sHelpAbout = Max Then
        sHelpAbout = Min
        For temp = 1 To 1695 - 255
            picHelpAbout.Height = picHelpAbout.Height - 1
            DoEvents
        Next '  TEMP '  TEMP
        lblAbuthelp.BackColor = &H8000000F
     ElseIf sHelpAbout = Min Then 'NOT SHELPABOUT...
        sHelpAbout = Max
        If sPicScan = Max Then
            lblFileScan_Click
        End If
        If SpicOther = Max Then
            lblOther_Click
        End If
        For temp = 1 To 1695 - 255
            picHelpAbout.Height = picHelpAbout.Height + 1
            DoEvents
        Next '  TEMP '  TEMP
        lblAbuthelp.BackColor = &HE0E0E0
    End If

End Sub

Private Sub lblBug_Click()

    Call ShellExecute(Me.hWnd, "Open", "mailto:cyber_chris235@gmx.net?subject=Bug in " & AV.AVname, vbNullString, "c:\", 1)
    MsgBox LoadResString(128) '"Thank you for your help!"

End Sub

Private Sub lblCFP_Click()

    picPathsearch_Click

End Sub

Private Sub lblffs_Click()

    Call picFastSearchx_Click(0)

End Sub

Private Sub lblFileScan_Click()

  Dim temp As Long

    If sPicScan = Max Then
        sPicScan = Min
        For temp = 1 To 2415 - 255
            picScan.Height = picScan.Height - 1
            picOther.Top = picScan.Top + picScan.Height + 20
            picHelpAbout.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        lblFileScan.BackColor = &H8000000F
     ElseIf sPicScan = Min Then 'NOT SPICSCAN...
        sPicScan = Max
        If sHelpAbout = Max Then
            lblAbuthelp_Click
        End If
        If SpicOther = Max Then
            lblOther_Click
        End If
        For temp = 1 To 2415 - 255
            picScan.Height = picScan.Height + 1
            picOther.Top = picScan.Top + picScan.Height + 20
            picHelpAbout.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        lblFileScan.BackColor = &HE0E0E0
    End If

End Sub

Private Sub lblHelp_Click()

    picHelp_Click

End Sub

Private Sub lblLogFile_Click()
frmLog.Show
End Sub

Private Sub lblOther_Click()

  Dim temp As Long

    If SpicOther = Max Then
        SpicOther = Min
        For temp = 1 To 1695 - 255
            picOther.Height = picOther.Height - 1
            picHelpAbout.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        lblOther.BackColor = &H8000000F
     ElseIf SpicOther = Min Then 'NOT SPICOTHER...
        SpicOther = Max
        If sHelpAbout = Max Then
            lblAbuthelp_Click
        End If
        If sPicScan = Max Then
            lblFileScan_Click
        End If
        For temp = 1 To 1695 - 255
            picOther.Height = picOther.Height + 1
            picHelpAbout.Top = picOther.Top + picOther.Height + 20
            DoEvents
        Next '  TEMP '  TEMP
        lblOther.BackColor = &HE0E0E0
    End If

End Sub

Private Sub lblSecured_Click()

    picSec_Click

End Sub

Private Sub lblSif_Click()

    picFileSearch_Click

End Sub

Private Sub lblText_Click(Index As Integer)

    On Error Resume Next
    If Index = 6 Then
        If lblText(7).Caption = "OFF" Then
            frmTray.Show , Me
         Else 'NOT LBLTEXT(7).CAPTION...
            Unload frmTray
        End If
     ElseIf Index = 9 Then 'NOT INDEX...
        If lblText(8).Caption = "OFF" Then
            SetKeyValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", AV.AVname, App.path & "\" & App.EXEName & ".exe /T", 1
            lblText(8).Caption = "ON"
            SaveSetting AV.AVname, "Settings", "Startup", "ON"
         Else 'NOT LBLTEXT(8).CAPTION...
            DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", AV.AVname
            lblText(8).Caption = "OFF"
            SaveSetting AV.AVname, "Settings", "Startup", "OFF"
        End If
     ElseIf Index = 17 Then 'NOT INDEX...
        If lblText(18).Caption = "OFF" Then
            lblText(18).Caption = "ON"
            SaveSetting AV.AVname, "Settings", "LogFile", "ON"
         Else 'NOT LBLTEXT(18).CAPTION...
            lblText(18).Caption = "OFF"
            SaveSetting AV.AVname, "Settings", "LogFile", "OFF"
        End If
     ElseIf Index = 15 Then 'NOT INDEX...
        frmSecFiles.Show , Me
     ElseIf Index = 19 Then
        If lblText(19).Caption = "ON" Then
           lblText(19).Caption = "OFF"
           Timer1.Enabled = False
        Else
            lblText(19).Caption = "ON"
            Timer1.Enabled = True
        End If
    End If
    On Error GoTo 0

End Sub

Private Sub lblupdate_Click()

    picUpdate_Click

End Sub

Private Sub mnuExit_Click()

    End

End Sub

Private Sub mnuLMainWindow_Click()

    On Error Resume Next
    Load frmMain
    BuildUI
    On Error GoTo 0

End Sub

Private Sub mnuSF_Click()

    ShowFileSearch

End Sub

Private Sub picAbout_Click()

    frmAbout.Show , Me

End Sub

Private Sub picFastSearchx_Click(Index As Integer)

    X.ControlToSetNewParent = Picture1
    Debug.Print X.ShowOpen(Me.hWnd)

End Sub

Private Sub picFileSearch_Click()

    Call ShowFileSearch

End Sub

Private Sub picHelp_Click()

    frmHelpsystem.Show , Me

End Sub

Private Sub picPathsearch_Click()

    Checkfolder

End Sub

Private Sub picSec_Click()

    frmSecFiles.Show

End Sub

Private Sub picUpdate_Click()

    frmUpdate.Show , Me

End Sub

Public Sub ShowFileSearch()

  Dim strFilename As String

    On Error Resume Next
    strFilename = (ShowOpenDlg(Me, , LoadResString(129) & "|*.*", , "Scan File")) 'All files
    Debug.Print Len(strFilename)
    If Len(strFilename) > 3 Then 'avoids Bug on Cancel
        If FileLen(strFilename) <> 0 Then
            CheckFile (strFilename)
        End If
    End If
    On Error GoTo 0

End Sub

Private Sub Timer1_Timer()
Dim buffer, ValidData As String
Dim c As Collection
Dim currentlocation As String
    On Error Resume Next

    Timer1.Enabled = False
    For Each iewindow In currentwindows
        DoEvents
        If iewindow.Busy Then
            GoTo busysignal
        End If
        currentlocation = iewindow.LocationURL
        ValidData = InStr(1, buffer, iewindow.LocationName & "|" & iewindow.LocationURL & "|")
        If ValidData = 0 Then
            If Mid$(currentlocation, 1, 7) = "file://" Then
                 currentlocation = Replace(currentlocation, "file:///", "")
                 currentlocation = Replace(currentlocation, "%20", " ")
                 currentlocation = Replace(currentlocation, "/", "\")
                   FullPathSearch currentlocation, c
                   Debug.Print currentlocation
            End If
        End If
busysignal:
        
    Next
    Timer1.Enabled = True
    On Error GoTo 0
End Sub

Private Sub X_FileChanged(ByVal FileName As String)

    lblFileName.Caption = Mid$(FileName, InStrRev(FileName, "\") + 1)
    lblText(12).Caption = FileLen(FileName) & " Bytes"
    lblText(14).Caption = CalcCRC(FileName)
    CurrentFile = FileName

End Sub

