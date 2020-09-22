Attribute VB_Name = "modLanguage"
Option Explicit
Private LanguagePack()               As String    'The Signatures will be loaded into this array
'Private SignStr()            As String
'Private SignVirusStringType() As String * 1
'Private SignVirusName()      As String

Public Sub BuildTranslation()

  'This builds the Signature - Array
  
  Dim sIn        As String
  Dim swords()   As String
  Dim X          As Long
  Dim Y          As Long

  Dim Data()     As String
  If AV.Language.Lanugage <> "" Then
    sIn = FileText(App.path & "\language\" & AV.Language.Lanugage & ".lng")
    swords = Split(sIn, vbCrLf)
    ReDim Preserve swords(UBound(swords) - 1)
    sIn = ""
    For X = LBound(swords) To UBound(swords)
        ReDim Preserve LanguagePack(0 To X) As String
        'Data = Split(swords(X) & ":" & ":", ":")
        LanguagePack(X) = swords(X)
        
    Next X
    ReDim Preserve LanguagePack(0 To X + 1) As String
    If Mid(LanguagePack(X - 1), 1, 13) = "##Translator#" Then
        AV.Language.Translator = Mid(LanguagePack(X - 1), 14, Len(LanguagePack(X - 1)) - 14)
    Else
        MsgBox "An error hase occoured while loading the Languagepack!"
    End If
    End If
End Sub

Public Sub TranslateLabel(TControl As Control, _
                          ByVal Id As Long)
On Error GoTo err:

    TControl.Caption = Mid(LanguagePack(Id - 101), 1, Len(LanguagePack(Id - 101)) - 1)
Exit Sub
err:
    TControl.Caption = LoadResString(Id)
End Sub

