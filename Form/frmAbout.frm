VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "About"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picAward 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   735
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   10
      ToolTipText     =   "Planet Source Code Superior Coding Contest Winner"
      Top             =   1200
      Width           =   855
   End
   Begin CCAntivir2004.DMSXpButton cmdExit 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OK"
      ForeColor       =   -2147483630
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   3
      Text            =   "frmAbout.frx":22A2
      Top             =   1200
      Width           =   4935
   End
   Begin CCAntivir2004.Hyperlink lblthanks2 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      ForeColorIdle   =   16711680
      ForeColorMouse  =   255
      BackColor       =   16777215
      Caption         =   "Paul"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CCAntivir2004.Hyperlink lblThanks 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      ForeColorIdle   =   16711680
      ForeColorMouse  =   255
      BackColor       =   16777215
      Caption         =   "Patabugen"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CCAntivir2004.Hyperlink lbllCopyright 
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   3840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
      ForeColorIdle   =   16711680
      ForeColorMouse  =   255
      BackColor       =   16777215
      Caption         =   "cyber_chris235@gmx.net"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin CCAntivir2004.Hyperlink lblThanks3 
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      ForeColorIdle   =   16711680
      ForeColorMouse  =   255
      BackColor       =   16777215
      Caption         =   "Shamarq Systems"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "© Copyright by Cyber Chris"
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
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   5655
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Open Source Antivirus Project"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Cyber Chris"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Anti Virus 2004"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    Text1.Text = FileText(App.path & "\document\data.txt")
    lblText(3).Caption = "Version: " & App.major & "." & App.minor & "." & App.Revision & " - " & _
        AV.Language.Lanugage & " ( Translated by: " & AV.Language.Translator & " )"
    Me.Caption = LoadResString(130)  '"About"

End Sub

Private Sub Form_Unload(cancel As Integer)

  Dim myArticleAddr As String

    If MsgBox("Would you please vote on PSC Website in case you like this program?", vbQuestion + vbYesNo, "Your vote will be very well appreciated ...") = vbYes Then
        myArticleAddr = "http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&txtCodeId=53497&optCodeRatingValue=5"
        Call ShellExecute(Me.hWnd, "Open", myArticleAddr, vbNullString, vbNullString, 1)
        MsgBox "Thank you very much. I really appreciate that :-) ", , "Thanks a million..."
    End If

End Sub

Private Sub lblCopyright_Click(Index As Integer)

    Call ShellExecute(Me.hWnd, "Open", "mailto:cyber_chris235@gmx.net", vbNullString, "c:\", 1)

End Sub

Private Sub lbllCopyright_Click()

    Call ShellExecute(Me.hWnd, "Open", "mailto:cyber_chris235@gmx.net", vbNullString, "c:\", 1)

End Sub

Private Sub lblthanks2_Click()


    Call ShellExecute(Me.hWnd, "Open", "mailto:wpsjr1@succeed.net", vbNullString, "c:\", 1)

End Sub

Private Sub lblThanks_Click()

    Call ShellExecute(Me.hWnd, "Open", "mailto:dude@patabugen.co.uk", vbNullString, "c:\", 1)

End Sub

Private Sub lblThanks3_Click()

    Call ShellExecute(Me.hWnd, "Open", "mailto:sharmaq@terra.com.br", vbNullString, "c:\", 1)

End Sub
