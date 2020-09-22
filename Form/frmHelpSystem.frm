VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmHelpsystem 
   Caption         =   "Help System"
   ClientHeight    =   7290
   ClientLeft      =   1785
   ClientTop       =   1470
   ClientWidth     =   11430
   LinkTopic       =   "Form10"
   ScaleHeight     =   7290
   ScaleWidth      =   11430
   Begin VB.ListBox List2 
      Height          =   6900
      IntegralHeight  =   0   'False
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2715
   End
   Begin VB.CommandButton cmdZurück 
      Height          =   315
      Left            =   5160
      Picture         =   "frmHelpSystem.frx":0000
      Style           =   1  'Grafisch
      TabIndex        =   10
      ToolTipText     =   "Zurück"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdVorwärts 
      Height          =   315
      Left            =   5580
      Picture         =   "frmHelpSystem.frx":038A
      Style           =   1  'Grafisch
      TabIndex        =   9
      ToolTipText     =   "Vorwärts"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdAuffrischen 
      Height          =   315
      Left            =   6420
      Picture         =   "frmHelpSystem.frx":0714
      Style           =   1  'Grafisch
      TabIndex        =   8
      ToolTipText     =   "Auffrischen"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdHome 
      Height          =   315
      Left            =   6840
      Picture         =   "frmHelpSystem.frx":0A9E
      Style           =   1  'Grafisch
      TabIndex        =   7
      ToolTipText     =   "Home"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdSuchen 
      Height          =   315
      Left            =   7260
      Picture         =   "frmHelpSystem.frx":0E28
      Style           =   1  'Grafisch
      TabIndex        =   6
      ToolTipText     =   "Finde auf dieser Seite"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdStop 
      Height          =   315
      Left            =   6000
      Picture         =   "frmHelpSystem.frx":11B2
      Style           =   1  'Grafisch
      TabIndex        =   5
      Tag             =   "Stop"
      ToolTipText     =   "Vorwärts"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdOptionen 
      Height          =   315
      Left            =   8100
      Picture         =   "frmHelpSystem.frx":153C
      Style           =   1  'Grafisch
      TabIndex        =   4
      ToolTipText     =   "Internet Optionen"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdErweiterte_Suche 
      Height          =   315
      Left            =   7680
      Picture         =   "frmHelpSystem.frx":18C6
      Style           =   1  'Grafisch
      TabIndex        =   3
      ToolTipText     =   "Erweiterte Suche"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton cmdDrucken_Vorschau 
      Height          =   315
      Left            =   8520
      Picture         =   "frmHelpSystem.frx":1C50
      Style           =   1  'Grafisch
      TabIndex        =   2
      ToolTipText     =   "Drucken Vorschau"
      Top             =   120
      Width           =   375
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Unten ausrichten
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   6975
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   16757
            MinWidth        =   16757
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3351
            MinWidth        =   3351
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6255
      Left            =   2760
      TabIndex        =   0
      Top             =   600
      Width           =   8655
      ExtentX         =   15266
      ExtentY         =   11033
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   6540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpSystem.frx":1FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpSystem.frx":23B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpSystem.frx":2780
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmHelpsystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAuffrischen_Click()

    On Error Resume Next
    WebBrowser1.Refresh
    On Error GoTo 0

End Sub

Private Sub cmdDrucken_Vorschau_Click()

    On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT
    On Error GoTo 0

End Sub

Private Sub cmdErweiterte_Suche_Click()

    On Error Resume Next
    WebBrowser1.GoSearch
    WebBrowser1.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DONTPROMPTUSER
    On Error GoTo 0

End Sub

Private Sub cmdHome_Click()


    On Error Resume Next
    WebBrowser1.GoHome
    On Error GoTo 0

End Sub

Private Sub cmdOptionen_Click()

    Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl"

End Sub


Private Sub cmdStop_Click()

    On Error Resume Next
    WebBrowser1.Stop
    On Error GoTo 0

End Sub

Private Sub cmdSuchen_Click()

    On Error Resume Next
    WebBrowser1.SetFocus
    SendKeys "^f", True
    On Error GoTo 0

End Sub

Private Sub cmdVorwärts_Click()


    On Error Resume Next
    WebBrowser1.GoForward
    On Error GoTo 0

End Sub

Private Sub cmdZurück_Click()


    On Error Resume Next
    WebBrowser1.GoBack
    On Error GoTo 0

End Sub

Private Sub Form_Load()

  Dim file_name As String

    On Error Resume Next
    LoadList List2, App.path & "\help\Index-tab.txt"
    file_name = App.path
    If Mid$(file_name, Len(file_name) - 1, 1) <> "\" Then
        file_name = file_name & "\"
    End If
    file_name = file_name & "\help\Node.txt"
    WebBrowser1.Navigate "about:blank"
    On Error GoTo 0

End Sub

Private Sub List2_Click()
  
  Dim sHTMFile As String

    With List2
        If .ListIndex >= 0 Then
            If InStr(.Text, vbTab) > 0 Then
                sHTMFile = Mid$(.Text, InStrRev(.Text, vbTab) + 1)
             Else 'NOT INSTR(.TEXT,...
                sHTMFile = .Text & ".htm"
            End If
        End If
        WebBrowser1.Navigate "file:\\" & App.path & "\Help\" & sHTMFile
    End With 'LIST2

End Sub


