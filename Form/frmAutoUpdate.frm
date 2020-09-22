VERSION 5.00
Begin VB.Form frmAutoUpdate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Update Reminder"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin CCAntivir2004.DMSXpButton cmdOK 
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   1920
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "&OK"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CCAntivir2004.DMSXpButton cmdNo 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "&No"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CCAntivir2004.DMSXpButton cmdYes 
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "&Yes"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.CheckBox opRemind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remind later:"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtRemindLater 
      Height          =   285
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "1"
      Top             =   1920
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4560
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "day(s)"
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   4
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Should I check for new updates?"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "days."
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Your signature file is quite old:"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmAutoUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub cmdNo_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    If opRemind.Value = 1 Then
        SaveSetting AV.AVname, "Settings", "RemindLater", DateAdd("d", txtRemindLater.Text, date)
    End If
    Unload Me

End Sub

Private Sub cmdYes_Click()

    frmUpdate.Show , Me
    Unload Me

End Sub

Private Sub Form_Load()

    lblText(1).Caption = DateDiff("d", frmMain.lblText(3).Caption, date)
    TranslateLabel lblText(2), 108
    TranslateLabel lblText(0), 109
    TranslateLabel lblText(3), 110
    TranslateLabel opRemind, 111
    TranslateLabel lblText(5), 108
    TranslateLabel cmdYes, 101
    TranslateLabel cmdNo, 102

End Sub

Private Sub txtRemindLater_KeyDown(KeyCode As Integer, _
                                   Shift As Integer)

  '48..57


End Sub

Private Sub txtRemindLater_KeyPress(KeyAscii As Integer)

  'Numbers only

    If KeyAscii - 48 > 9 Then
        KeyAscii = 0
    End If

End Sub


