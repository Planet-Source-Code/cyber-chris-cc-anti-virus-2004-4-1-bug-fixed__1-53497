VERSION 5.00
Begin VB.Form frmAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "CCAntivir 2004"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin CCAntivir2004.Hyperlink hpOnline 
      Height          =   225
      Left            =   5205
      TabIndex        =   13
      Top             =   2670
      Width           =   2115
      _extentx        =   3731
      _extenty        =   397
      forecoloridle   =   16711680
      forecolormouse  =   255
      backcolor       =   16777215
      caption         =   "More Virus Information Online"
   End
   Begin CCAntivir2004.DMSXpButton cmdSecure 
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   2160
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      font            =   "frmAlert.frx":0000
      caption         =   "&Secure"
      forecolor       =   -2147483642
      forehover       =   0
   End
   Begin CCAntivir2004.DMSXpButton cmdRemove 
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   2160
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      font            =   "frmAlert.frx":002C
      caption         =   "&Remove"
      forecolor       =   -2147483642
      forehover       =   0
   End
   Begin CCAntivir2004.DMSXpButton cmdIgnore 
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
      _extentx        =   2355
      _extenty        =   661
      font            =   "frmAlert.frx":0058
      caption         =   "&Ignore"
      forecolor       =   -2147483642
      forehover       =   0
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   585
      Left            =   6000
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   3
      Top             =   960
      Width           =   645
   End
   Begin VB.Line lline 
      Index           =   3
      X1              =   6840
      X2              =   6840
      Y1              =   2040
      Y2              =   480
   End
   Begin VB.Line lline 
      Index           =   2
      X1              =   480
      X2              =   6840
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line lline 
      Index           =   1
      X1              =   480
      X2              =   480
      Y1              =   480
      Y2              =   2040
   End
   Begin VB.Line lline 
      Index           =   0
      X1              =   480
      X2              =   6840
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   9
      Top             =   1680
      Width           =   4815
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   1800
      TabIndex        =   8
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "File size:"
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename:"
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   4815
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Virus found!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub BuildAlert()

    On Error Resume Next
    lblText(1).Caption = Virus.Filename
    lblText(1).ToolTipText = Virus.Filename & "  (" & FileLen(Virus.Filename) & " Bytes )"
    lblText(2).Caption = Virus.Reason
    lblText(8).Caption = FileLen(Virus.Filename) & " Bytes"
    If Virus.Type = Executable Then
        lblText(7).Caption = "Executable File"
    End If
    If Virus.Type = Script Then
        lblText(7).Caption = "Script"
    End If
    picIcon.Picture = LoadIcon(Large, Virus.Filename)
    On Error GoTo 0

End Sub

Private Sub cmdIgnore_Click()

    Log "Alert ignored: " & Virus.Reason, 2
    Unload Me

End Sub

Private Sub cmdRemove_Click()

    Log "File removed: " & Virus.Filename, 2
    RemoveFile (Virus.Filename)

End Sub

Private Sub cmdSecure_Click()

  Dim sXor As New clsSimpleXOR

    On Error Resume Next
    MsgBox "The File will be secured, that means everytime you want to start it, you'll get a prompt." & vbCrLf & _
     "This will avoid unwanted starts!", vbInformation + vbOKOnly
    sXor.EncryptFile Virus.Filename, Virus.Filename, AV.AVname
    Set sXor = Nothing
    MkDir App.path & "\Secure\"
    FileCopy Virus.Filename, App.path & "\Secure\" & Mid$(Virus.FileNameShort, 1, Len(Virus.FileNameShort) - 1) & ".secure"
    Kill Virus.Filename
    With frmSecFiles
        .Visible = False
        .Show
        SaveSetting AV.AVname, "Settings", "Quarintine", .flSec.ListCount
    End With 'frmSecFiles
    Unload frmSecFiles
    Log "File moved to quarintine: " & Virus.Filename, 2
    On Error GoTo 0

End Sub



Private Sub Form_Load()
Debug.Print Time
  Dim R1  As RECT
  Dim R2  As RECT
  Dim TPP As Integer

    TPP = Screen.TwipsPerPixelX
    Call SetRect(R1, Screen.Width / TPP, Screen.Height / TPP, Screen.Width / TPP, Screen.Height / TPP)
    Call SetRect(R2, 0, 0, Me.Width / TPP, Me.Height / TPP)
    Call DrawAnimatedRects(Me.hWnd, IDANI_CLOSE Or IDANI_CAPTION, R1, R2)
    BuildAlert
    KeepOnTop Me
    TranslateLabel hpOnline, 150
    TranslateLabel lblText(0), 103
    TranslateLabel lblText(3), 104
    TranslateLabel lblText(4), 105
    TranslateLabel lblText(6), 106
    TranslateLabel lblText(5), 107
    TranslateLabel cmdIgnore, 132
    TranslateLabel cmdRemove, 133
    TranslateLabel cmdSecure, 134
    DoEvents
    BeepAlert
End Sub

Private Sub BeepAlert()
    Beep 4000, 220
    Beep 3000, 200
    Beep 4000, 220
    Beep 3000, 200
End Sub

Private Sub hpOnline_Click()

  'http://www.viruslist.com/eng/viruslistfind.html?findTxt=code+red

    Call ShellExecute(Me.hWnd, "Open", "http://www.viruslist.com/eng/viruslistfind.html?findTxt=" & Replace(Virus.Reason, " ", "+"), vbNullString, "c:\", 1)

End Sub


