VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmLog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Log View"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9030
   StartUpPosition =   2  'Bildschirmmitte
   Begin CCAntivir2004.DMSXpButton btnSearch 
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   480
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
      Caption         =   "Search"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin CCAntivir2004.DMSXpButton BtnDone 
      Height          =   375
      Left            =   3960
      TabIndex        =   19
      Top             =   1200
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
      Caption         =   "Done"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin RichTextLib.RichTextBox tbOutput 
      Height          =   3495
      Left            =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2400
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   6165
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmLog.frx":0442
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
   Begin VB.TextBox tbQuery 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   360
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox tbData 
         Height          =   285
         Left            =   840
         TabIndex        =   14
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox cmbModule 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown-Liste
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox tbSocket 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   3
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox tbTimeTo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   5
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox tbTimeFrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "HH:mm"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   4
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   5
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox tbDateTo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox tbDateFrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   840
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFFFF&
         Caption         =   "( use '0' to list All Sockets )"
         Height          =   375
         Left            =   1800
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFFFF&
         Caption         =   "Data : "
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFFFF&
         Caption         =   " to "
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00FFFFFF&
         Caption         =   " to "
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFFFF&
         Caption         =   "Module : "
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFFFF&
         Caption         =   "Socket : "
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFFFF&
         Caption         =   "Time : "
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Rechts
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date : "
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Your Query :"
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
      Left            =   5640
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Const Log_file = "log.txt"

Public Sub Form_Load()
  
  Dim i As Integer

  FileCopy App.path & "/" & Log_file, App.path & "/server.bak"
  
  cmbModule.Clear
  cmbModule.AddItem ("ALL")
  
  For i = 1 To MAX_TYPES
    cmbModule.AddItem (logType(i))
  Next i
  cmbModule.ListIndex = 0
End Sub

Private Sub Form_Unload(cancel As Integer)

  On Error Resume Next
  Kill App.path & "/server.bak"
  If (err.Number <> 0) Then err.Clear

  Unload Me
End Sub

Public Sub WriteLog(st As String)

  st = Trim(st)
  Open App.path & "\" & Log_file For Append As #9
  Print #9, st
  Close #9
End Sub

Public Sub AddLog(typ As Integer, skt As Integer, st As String)


  Dim ns As String
    
  ns = Trim(CStr(skt))
  If (Len(ns) < 3) Then ns = Left("000", 3 - Len(ns)) & ns
  
  st = Trim(st)
  st = ns & ": [" & logType(typ) & "] " & st
  st = Format(Now, "MM/DD/YYYY HH:MM") & " " & st
  

  Call WriteLog(st)
End Sub

Private Sub QueryLog()

  Dim srchSdate As Date
  Dim srchEdate As Date
  Dim srchSock As Integer
  Dim srchMod As String
  Dim srchData As String
  Dim lineDate As Date
  Dim lineSock As Integer
  Dim lineMod As String
  Dim lineData As String
  Dim st, temp As String
  Dim CRLF As String
  Dim foundLin As Boolean
  Dim ctr As Integer
  Dim i As Integer
  Dim e As Integer
  
  CRLF = Chr(13) & Chr(10)
 'CREATE QUERY
  If (tbTimeFrom.Text = "") Then
    st = "00:00"
  Else
    st = tbTimeFrom.Text
  End If
  
  If (tbDateFrom.Text = "") Then
    temp = "01/01/2000"
  Else
    temp = tbDateFrom.Text
  End If
  
  temp = temp & " " & st
  srchSdate = CDate(Format(temp, "MM/DD/YYYY HH:MM"))
  
  If (tbTimeTo.Text = "") Then
    If (tbDateTo.Text = "") Then
      st = Format(Now, "HH:MM")
    Else
      st = "23:59"
    End If
  Else
    st = tbTimeTo.Text
  End If
  
  If (tbDateTo.Text = "") Then
    temp = Format(Now, "MM/DD/YYYY")
  Else
    temp = tbDateTo.Text
  End If
  
  temp = temp & " " & st
  srchEdate = CDate(Format(temp, "MM/DD/YYYY HH:MM"))
  
  If (tbSocket.Text = "") Then
    srchSock = 0
  Else
    srchSock = Trim(CStr(tbSocket.Text))
  End If
  
  srchMod = cmbModule.List(cmbModule.ListIndex)
  srchData = Trim(tbData.Text)
  
 'DISPLAY QUERY
  tbQuery.Text = ""
  tbQuery.Text = tbQuery.Text & "Start Date: " & Format(srchSdate, "MM/DD/YYYY HH:MM") & CRLF
  tbQuery.Text = tbQuery.Text & "  End Date: " & Format(srchEdate, "MM/DD/YYYY HH:MM") & CRLF
  
  If (srchock = 0) Then
    tbQuery.Text = tbQuery.Text & "   Sockets: ALL" & CRLF
  Else
    tbQuery.Text = tbQuery.Text & "    Socket: " & Trim(CStr(srchSock)) & CRLF
  End If
  
  tbQuery.Text = tbQuery.Text & "    Module: " & srchMod & CRLF
  tbQuery.Text = tbQuery.Text & "      Data: " & srchData & CRLF
  tbQuery.Visible = True
  tbQuery.Refresh
  
  tbOutput.Text = ""
  tbOutput.Visible = False
  foundline = False
  ctr = 0
  
 'DISPLAY QUERY RESULTS
  Open App.path & "/server.bak" For Input As #8
  Do While (Not EOF(8))
    st = ""
    temp = ""
    Do While (temp <> Chr(10)) And (Not EOF(8))
      temp = Input(1, #8)
      st = st & temp
    Loop
    
    If (InStr(1, st, Chr(13)) > 0) Then st = Mid(st, 1, InStr(1, st, Chr(13)) - 1)
    
    If (Trim(CStr(Mid(st, 1, 2))) > 0) Then
      lineDate = CDate(Format(Mid(st, 1, 16), "MM/DD/YYYY HH:MM"))
    Else
      lineDate = CDate(Format("01/01/1990 01:00", "MM/DD/YYYY HH:MM"))
    End If
    
    lineSock = Trim(CStr(Mid(st, 18, 3)))
    lineMod = Mid(st, 24, MAX_LEN)
    lineData = Trim(Mid(st, 26 + MAX_LEN, Len(st)))
    
    If (lineDate >= srchSdate) And (lineDate <= srchEdate) And ((lineSock = srchSock) Or (srchSock = 0)) And ((lineMod = srchMod) Or (srchMod = "ALL")) And (InStr(1, lineData, srchData) > 0) Then
      tbOutput.Text = tbOutput.Text & st & CRLF
      foundline = True
    End If
    
    DoEvents
  Loop
  Close #8
  If (Not foundline) Then tbOutput.Text = "No Lines match your Query."
  tbOutput.Visible = True
End Sub

Private Sub btnDone_Click()
  Call Form_Unload(0)
End Sub

Private Sub btnSearch_Click()
  Call QueryLog
End Sub

Private Sub tbDateFrom_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) And (Chr(KeyAscii) <> "/") Then KeyAscii = 0
End Sub

Private Sub tbDateFrom_GotFocus()
  SendKeys "{home}+{end}"
End Sub

Private Sub tbDateTo_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) And (Chr(KeyAscii) <> "/") Then KeyAscii = 0
End Sub

Private Sub tbDateTo_GotFocus()
  If (tbDateFrom.Text <> "") And (tbDateTo.Text = "") Then tbDateTo.Text = tbDateFrom.Text
  SendKeys "{home}+{end}"
End Sub

Private Sub tbTimeFrom_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) And (Chr(KeyAscii) <> ":") Then KeyAscii = 0
End Sub

Private Sub tbTimeFrom_GotFocus()
  SendKeys "{home}+{end}"
End Sub

Private Sub tbTimeTo_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) And (Chr(KeyAscii) <> ":") Then KeyAscii = 0
End Sub

Private Sub tbTimeTo_GotFocus()
  If (tbDateFrom.Text <> "") And (tbDateTo.Text = "") Then tbTimeTo.Text = tbTimeFrom.Text
  SendKeys "{home}+{end}"
End Sub

Private Sub tbSocket_KeyPress(KeyAscii As Integer)
  If Not IsNumeric(Chr(KeyAscii)) And (KeyAscii <> 8) Then KeyAscii = 0
End Sub

Private Sub tbSocket_GotFocus()
  SendKeys "{home}+{end}"
End Sub

Private Sub tbData_GotFocus()
  SendKeys "{home}+{end}"
End Sub


