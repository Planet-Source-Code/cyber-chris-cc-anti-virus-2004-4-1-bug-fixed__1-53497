VERSION 5.00
Begin VB.Form frmTray 
   BorderStyle     =   4  'Festes Werkzeugfenster
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   555
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Basket 
      Height          =   435
      Left            =   0
      OLEDropMode     =   1  'Manuell
      Picture         =   "frmTray.frx":0000
      ToolTipText     =   "Drag & Drop here the files you want to have checked! or doubleklick to Start main Wndow!"
      Top             =   0
      Width           =   525
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub Basket_DblClick()

    AV.Runmode = Normal
    frmMain.Show
    Unload Me

End Sub

Private Sub Basket_OLEDragDrop(Data As DataObject, _
                               Effect As Long, _
                               Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

  Dim file As Variant
  Dim k    As Long

    On Error GoTo droperror
    k = 0
    For Each file In Data.Files
        If LenB(Dir(file)) > 0 And Not vbDirectory Then
            CheckFile (file)
            k = k + 1
        End If
    Next
    TranslateLabel Me.lblStatus, 123
    Me.lblStatus.Caption = lblStatus.Caption & " " & k

Exit Sub

droperror:

End Sub

Private Sub Form_Load()

    KeepOnTop Me
    TranslateLabel Me.lblStatus, 123
    Me.lblStatus.Caption = lblStatus.Caption & " " & 0
    frmMain.lblText(7).Caption = "ON"
    Me.Left = GetSetting(AV.AVname, "Settings", "TrayLeft", 0)
    Me.Top = GetSetting(AV.AVname, "Settings", "TrayTop", 0)

End Sub

Private Sub Form_Unload(cancel As Integer)

    frmMain.lblText(7).Caption = "OFF"
    SaveSetting AV.AVname, "Settings", "TrayLeft", Me.Left
    SaveSetting AV.AVname, "Settings", "TrayTop", Me.Top
    If AV.Runmode = TrayOnly Then
        End
    End If

End Sub


