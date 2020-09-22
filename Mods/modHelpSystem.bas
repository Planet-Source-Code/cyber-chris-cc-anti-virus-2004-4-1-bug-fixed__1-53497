Attribute VB_Name = "modList"
Option Explicit

' Benötigte API-Deklarationen
' INI lesen+schreiben
Private Declare Function OSGetPrivateProfileString Lib "KERNEL32" _
  Alias "GetPrivateProfileStringA" ( _
  ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpDefault As String, _
  ByVal lpReturnedString As String, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileString Lib "KERNEL32" _
  Alias "WritePrivateProfileStringA" ( _
  ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpString As Any, _
  ByVal lpFileName As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096

' ListBox-Inhalt in INI-Datei speichern
Public Sub SaveList(oListBox As ListBox, ByVal FileName As String)
  Dim nCount As Long
    
  On Error Resume Next
  ' Löscht die alte Datei um sicher zu sein,
  ' dass jede extra Information die nicht
  ' benötigt wird gelöscht wird
  Kill FileName
  
  With oListBox
    nCount = 0
    
    ' alle ListBox-Einträge speichern
    Do
      nCount = nCount + 1
      WriteINIString "List", CStr(nCount), .List(nCount - 1), FileName
    Loop Until nCount = .ListCount
  End With
End Sub
' ListBox-Inhalt aus INI-Datei auslesen
Public Sub LoadList(oListBox As ListBox, ByVal FileName As String)
  Dim nCount As Long
  Dim sText As String
  
  nCount = 0
  
  ' Inhalt der Liste löschen
  With oListBox
    .Clear

    Do
      ' Eintrag aus INI-Liste lesen
      nCount = nCount + 1
      sText = GetINIString("List", CStr(nCount), "", FileName)
      
      ' Fügt den Eintrag zur Liste hinzu
      If sText <> "" Then
        If InStr(sText, "|") > 0 Then sText = Replace(sText, "|", String$(10, vbTab))
        .AddItem sText
      End If
    Loop Until sText = ""
  End With
End Sub

' INI-Eintrag lesen
Private Function GetINIString(ByVal szSection As String, _
  ByVal szEntry As Variant, _
  ByVal szDefault As String, _
  ByVal szFileName As String) As String

  Dim szTmp As String
  Dim nRet  As Long

  If (IsNull(szEntry)) Then
    szTmp = String$(nBUFSIZEINIALL, 0)
    nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, _
      szTmp, nBUFSIZEINIALL, szFileName)
  Else
    szTmp = String$(nBUFSIZEINI, 0)
    nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), _
      szDefault, szTmp, nBUFSIZEINI, szFileName)
  End If
  
  GetINIString = Left$(szTmp, nRet)
End Function

' INI-Eintrag speichern
Private Sub WriteINIString(ByVal szSection As String, _
  ByVal szEntry As Variant, _
  ByVal vValue As Variant, _
  ByVal szFileName As String)

  Dim nRet As Long

  If (IsNull(szEntry)) Then
    nRet = OSWritePrivateProfileString(szSection, 0&, 0&, _
      szFileName)
  ElseIf (IsNull(vValue)) Then
    nRet = OSWritePrivateProfileString(szSection, CStr(szEntry), _
      0&, szFileName)
  Else
    nRet = OSWritePrivateProfileString(szSection, CStr(szEntry), _
      CStr(vValue), szFileName)
  End If
End Sub

' Prüfen, ob Datei existiert
Public Function FileExists(ByVal sFile As String) As Boolean
  On Error Resume Next
  FileExists = (Dir$(sFile) <> "")
  On Error GoTo 0
End Function
