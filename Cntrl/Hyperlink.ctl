VERSION 5.00
Begin VB.UserControl Hyperlink 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MouseIcon       =   "Hyperlink.ctx":0000
   MousePointer    =   99  'Benutzerdefiniert
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3990
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4410
      Top             =   0
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   225
      Left            =   0
      Shape           =   4  'Gerundetes Rechteck
      Top             =   0
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      MouseIcon       =   "Hyperlink.ctx":08CA
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   0
      Top             =   0
      Width           =   990
   End
End
Attribute VB_Name = "Hyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function SetCapture Lib "User32" (ByVal hWnd As Long) As Long
Private Declare Function LockWindowUpdate Lib "User32" (ByVal hWndLock As Long) As Integer

Private Type SepRGB
    Red As Long
    Green As Long
    Blue As Long
End Type

Dim ForeIdle As Long
Dim ForeMouse As Long
Dim ShowRec As Boolean
Private FadeOut As Boolean
Private FadeOut2 As Boolean
Dim SC As Boolean
Dim FC As Boolean
Dim PendingCaption As String
Event Click()
Event Change()

Sub DoMouseActions(MouseIn As Boolean)
If FC = True Then Exit Sub

If MouseIn = True Then
    If ShowRec = True Then If Shape1.Visible = False Then Shape1.Visible = True
    'Label1.ForeColor = ForeMouse
    FadeOut = False
    Timer1.Enabled = True
    If SC = True Then
        Image1.Visible = False
        Image2.Visible = True
    End If
        Else
            Shape1.Visible = False
            'Label1.ForeColor = ForeIdle
            FadeOut = True
            Timer1.Enabled = True
            If SC = True Then
                Image1.Visible = True
                Image2.Visible = False
            End If
End If
End Sub
Sub DoMouseOutFade(ReturnImed As Boolean)
Label1.ForeColor = ForeMouse
FadeOut = True
Timer1.Enabled = True
Timer1_Timer
If ReturnImed = False Then
    Do Until Timer1.Enabled = False: DoEvents: Loop
End If
End Sub
Property Let ShowRectangle(NewValue As Boolean)
ShowRec = NewValue
End Property
Property Get ShowRectangle() As Boolean
ShowRectangle = ShowRec
End Property
Property Let ForeColorIdle(NewColor As OLE_COLOR)
If FC = True Then
    NewColor = 0
End If

Label1.ForeColor = NewColor
ForeIdle = NewColor
End Property

Property Get ForeColorIdle() As OLE_COLOR
ForeColorIdle = ForeIdle
End Property

Property Let ForeColorMouse(NewValue As OLE_COLOR)
ForeMouse = NewValue
End Property

Property Get ForeColorMouse() As OLE_COLOR
ForeColorMouse = ForeMouse
End Property

Property Let BackColor(NewValue As OLE_COLOR)
If FC = True Then
    NewValue = vbWhite
End If

Label1.BackColor = NewValue
UserControl.BackColor = NewValue
End Property

Property Get BackColor() As OLE_COLOR
BackColor = Label1.BackColor
End Property

Private Sub Label1_Change()
RaiseEvent Change
End Sub

Private Sub Label1_Click()
RaiseEvent Click
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UserControl_MouseMove Button, Shift, X, Y
End Sub


Private Sub Timer1_Timer()
Dim CurRGB As SepRGB
Dim InBy As SepRGB
Dim ForeIdleRGB As SepRGB
Dim ForeMouseRGB As SepRGB
On Error Resume Next

ForeMouseRGB = GetRGB(ForeMouse)
ForeIdleRGB = GetRGB(ForeIdle)

If FadeOut = True Then
    CurRGB = GetRGB(Label1.ForeColor)
    
    With InBy
        .Red = Abs(CurRGB.Red - ForeIdleRGB.Red)
        .Green = Abs(CurRGB.Green - ForeIdleRGB.Green)
        .Blue = Abs(CurRGB.Blue - ForeIdleRGB.Blue)
        
        .Red = .Red / 7
        .Green = .Green / 7
        .Blue = .Blue / 7
        
        If .Red = 0 And .Green = 0 And .Blue = 0 Then
            Timer1.Enabled = False
            LockWindowUpdate UserControl.hWnd
            Label1.ForeColor = ForeIdle
            LockWindowUpdate 0
            Exit Sub
        End If
    End With
    
    With CurRGB
        If ForeIdleRGB.Red <> .Red Then If ForeIdleRGB.Red < .Red Then .Red = .Red - InBy.Red Else .Red = .Red + InBy.Red
        If ForeIdleRGB.Green <> .Green Then If ForeIdleRGB.Green < .Green Then .Green = .Green - InBy.Green Else .Green = .Green + InBy.Green
        If ForeIdleRGB.Blue <> .Blue Then If ForeIdleRGB.Blue < .Blue Then .Blue = .Blue - InBy.Blue Else .Blue = .Blue + InBy.Blue
        LockWindowUpdate UserControl.hWnd
        Label1.ForeColor = RGB(.Red, .Green, .Blue)
        LockWindowUpdate 0
    End With
        Else
            CurRGB = GetRGB(Label1.ForeColor)

            With InBy
                .Red = Abs(CurRGB.Red - ForeMouseRGB.Red)
                .Green = Abs(CurRGB.Green - ForeMouseRGB.Green)
                .Blue = Abs(CurRGB.Blue - ForeMouseRGB.Blue)
                
                .Red = .Red / 4
                .Green = .Green / 4
                .Blue = .Blue / 4
                
                If .Red = 0 And .Green = 0 And .Blue = 0 Then
                    Timer1.Enabled = False
                    LockWindowUpdate UserControl.hWnd
                    Label1.ForeColor = ForeMouse
                    LockWindowUpdate 0
                    Exit Sub
                End If
            End With
            
            With CurRGB
                If ForeMouseRGB.Red <> .Red Then If ForeMouseRGB.Red < .Red Then .Red = .Red - InBy.Red Else .Red = .Red + InBy.Red
                If ForeMouseRGB.Green <> .Green Then If ForeMouseRGB.Green < .Green Then .Green = .Green - InBy.Green Else .Green = .Green + InBy.Green
                If ForeMouseRGB.Blue <> .Blue Then If ForeMouseRGB.Blue < .Blue Then .Blue = .Blue - InBy.Blue Else .Blue = .Blue + InBy.Blue
                LockWindowUpdate UserControl.hWnd
                Label1.ForeColor = RGB(.Red, .Green, .Blue)
                LockWindowUpdate 0
            End With
End If
End Sub

Private Sub Timer2_Timer()
Dim SepRGB As SepRGB
On Error Resume Next

If FadeOut2 = True Then
    SepRGB = GetRGB(Label1.ForeColor)
    With SepRGB
        .Blue = .Blue + 10
        .Red = .Red + 20
        .Green = .Green + 30
        Label1.ForeColor = RGB(.Red, .Green, .Blue)
        If Err Or Label1.ForeColor = vbWhite Then
            Label1.ForeColor = vbWhite
            FadeOut2 = False
            Label1.Caption = PendingCaption
        End If
    End With
        Else
            SepRGB = GetRGB(Label1.ForeColor)
            With SepRGB
                .Blue = .Blue - 10
                .Red = .Red - 20
                .Green = .Green - 30
                Label1.ForeColor = RGB(.Red, .Green, .Blue)
                If Err Or Label1.ForeColor = 0 Then
                    Label1.ForeColor = 0
                    Timer2.Enabled = False
                End If
            End With
End If
End Sub

Private Sub UserControl_Click()
RaiseEvent Click
End Sub
Function TextBoxHWND() As Long
TextBoxHWND = Text1.hWnd
End Function

Private Sub UserControl_InitProperties()
ForeColorIdle = 0
ForeColorMouse = &H80FF&
BackColor = vbButtonFace
Caption = Ambient.DisplayName
Set Font = Parent.Font
Enabled = True
Alignment = vbLeftJustify
ShowRectangle = False
TransparentBack = False
Speed = 50
ShowCarrot = False
FadeChange = False
End Sub

Private Function GetRGB(ByVal LongValue As Long) As SepRGB
LongValue = Abs(LongValue)
GetRGB.Red = LongValue And 255
GetRGB.Green = (LongValue \ 256) And 255
GetRGB.Blue = (LongValue \ 65536) And 255
End Function



Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static LastOne As Long

LastOne = SetCapture(hWnd)
If LastOne = hWnd Then LastOne = 0

If X < 0 Or X > UserControl.Width Or Y < 0 Or Y > UserControl.Height Then
    DoMouseActions False
    If LastOne = 0 Then
        ReleaseCapture
            Else
                SetCapture LastOne
    End If
        Else
            DoMouseActions True
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent Click
UserControl_MouseMove 0, 0, 1, 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
ForeColorIdle = PropBag.ReadProperty("ForeColorIdle", 0)
ForeColorMouse = PropBag.ReadProperty("ForeColorMouse", &H80FF&)
BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
Set Font = PropBag.ReadProperty("Font", Parent.Font)
Enabled = PropBag.ReadProperty("Enabled", True)
Alignment = PropBag.ReadProperty("Alignment", vbLeftJustify)
ShowRectangle = PropBag.ReadProperty("ShowRec", False)
Speed = PropBag.ReadProperty("Speed", 50)
ShowCarrot = PropBag.ReadProperty("ShowCarrot", False)
FadeChange = PropBag.ReadProperty("FC", False)
End Sub
Property Let Alignment(NewAlignment As AlignmentConstants)
Label1.Alignment = NewAlignment

If NewAlignment = vbRightJustify Then
    Shape1.Left = ScaleWidth - Shape1.Width
    Label1.Left = ScaleWidth - Label1.Width
    ElseIf NewAlignment = vbLeftJustify Then
        Shape1.Left = 0
    ElseIf NewAlignment = vbCenter Then
        Shape1.Left = 0
        Shape1.Width = Width
        Label1.AutoSize = False
        Label1.Width = Width
End If

End Property
Property Get Alignment() As AlignmentConstants
Alignment = Label1.Alignment
End Property
Private Sub UserControl_Resize()
If Alignment = vbCenter Then
    Label1.AutoSize = False
    Shape1.Width = Width
    Label1.Width = Width
        ElseIf Alignment = vbRightJustify Then
            Label1.Left = ScaleWidth - Label1.Width
        Else
            Label1.AutoSize = True
            Shape1.Width = Label1.Width + TextWidth(" ")
End If

Label1.Height = Height
If Label1.Alignment = vbCenter Then Shape1.Width = Width
Shape1.Height = Height


End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

PropBag.WriteProperty "ForeColorIdle", ForeIdle, 0
PropBag.WriteProperty "ForeColorMouse", ForeMouse, &H80FF&
PropBag.WriteProperty "BackColor", Label1.BackColor, vbButtonFace
If FC = True Then
    PropBag.WriteProperty "Caption", Label1.Caption, " " & Ambient.DisplayName
        Else
            PropBag.WriteProperty "Caption", Mid(Label1.Caption, 1, Len(Label1.Caption)), Ambient.DisplayName
End If
PropBag.WriteProperty "Font", Label1.Font, Parent.Font
PropBag.WriteProperty "Enabled", Label1.Enabled, True
PropBag.WriteProperty "Alignment", Label1.Alignment, vbLeftJustify
PropBag.WriteProperty "ShowRec", ShowRec, False
PropBag.WriteProperty "Speed", Timer1.Interval, 50
PropBag.WriteProperty "ShowCarrot", SC, False
PropBag.WriteProperty "FC", FC, False
End Sub

Property Get Caption() As String
If FC = True Then
    Caption = Label1.Caption
        Else
            Caption = Mid(Label1.Caption, 1, Len(Label1.Caption))
End If
End Property

Property Let Caption(NewValue As String)
If Label1.Caption = NewValue Then Exit Property

If FC = True Then
    PendingCaption = NewValue
    Timer2.Enabled = True
    FadeOut2 = True
    Exit Property
End If

Label1.Caption = NewValue
Shape1.Width = Label1.Width + TextWidth(" ")


If Label1.Alignment = vbCenter Then
    Label1.AutoSize = False
    Shape1.Width = Width
End If

If Label1.Alignment = vbRightJustify Then
    Shape1.Left = ScaleWidth - Shape1.Width
    ElseIf Label1.Alignment = vbCenter Then
        Shape1.Left = 0
End If
End Property

Property Get Font() As IFontDisp
Set Font = Label1.Font
End Property

Property Set Font(NewValue As IFontDisp)
Set Label1.Font = NewValue
Set UserControl.Font = NewValue
End Property

Property Let Enabled(NewValue As Boolean)
Label1.Enabled = NewValue
End Property

Property Get Enabled() As Boolean
Enabled = Label1.Enabled
End Property

Property Let Speed(NewValue As Integer)
On Error GoTo InvalidProp

Timer1.Interval = NewValue

Exit Property
InvalidProp:
MsgBox "Invalid property value", vbCritical
End Property

Property Get Speed() As Integer
Speed = Timer1.Interval
End Property

Property Let ShowCarrot(Value As Boolean)
SC = Value

If Value = False Then

    Label1.Left = 0
        Else
            Image1.Visible = True
            Label1.Left = Image2.Width
            ShowRectangle = False
End If
End Property

Property Get ShowCarrot() As Boolean
ShowCarrot = SC
End Property

Property Let FadeChange(NewValue As Boolean)
If NewValue = True Then
    ShowCarrot = False
    ForeColorIdle = 0
    BackColor = vbWhite
End If
FC = NewValue
End Property
Property Get FadeChange() As Boolean
FadeChange = FC
End Property
