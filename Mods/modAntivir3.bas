Attribute VB_Name = "modAntivir3"
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit
Public Function InStrFast(ByVal StringCheck As String, _
                         ByVal StringMatch As String, _
                Optional ByVal Start As Long = -1, _
                Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Long

  Dim lCheckLen  As Long   'length of stringcheck
  Dim lPosition  As Long   'current position

        lCheckLen = Len(StringCheck) 'length of stringcheck
        If (lCheckLen > 0) And (Len(StringMatch) > 0) Then
            If Start < 1 Then Start = lCheckLen
            If Start <= lCheckLen Then 'sanity check
                'loop/look backwards thru string
                Do While (Start > 0) And (lPosition = 0)
                    lPosition = InStr(Start, StringCheck, StringMatch, Compare)
                    If lPosition > Start Then lPosition = 0 'ignore positions greater than start
                    Start = Start - 1                       'de-increment position
                Loop
                InStrFast = lPosition
            End If
        End If
        
End Function
Public Sub CheckDLL()

    If FileExist(Environ$("Windir") & "\system32\unzip.dll") = False Then
        If MsgBox("The required dll 'unzip.dll' isn't present on this System!" & vbCrLf & "That means, that Zip Archive scanning isn't possible untill you download the file!" & vbCrLf & _
         "Please click yes to get further information", vbCritical + vbYesNo) = vbYes Then
            If MsgBox("The unzip.dll is a required dll which was developed by  the InfoZip group:" & vbCrLf & _
             "InfoZip http://www.cdrom.com/pub/infozip/" & vbCrLf & vbCrLf & "Should the file be downloaded now?", vbYesNo + vbQuestion) = vbYes Then
                modUpdate.DownloadFile "http://www.home.r-hs.de/philippinen/antivirus/dll/unzip.dll", Environ$("Windir") & "\system32\unzip.dll"
            End If
        End If
    End If

End Sub

Public Sub DelTree(sFolder As String)

  Dim sCurrFile As String

    sCurrFile = Dir(sFolder & "\*.*", vbDirectory)
    Do While Len(sCurrFile) > 0
        If sCurrFile <> "." And sCurrFile <> ".." Then
            If (GetAttr(sFolder & "\" & sCurrFile) And vbDirectory) = vbDirectory Then
                Call DelTree(sFolder & "\" & sCurrFile)
                sCurrFile = Dir(sFolder & "\*.*", vbDirectory)
             Else 'NOT (GETATTR(SFOLDER...
                Kill sFolder & "\" & sCurrFile
                sCurrFile = Dir
            End If
         Else 'NOT SCURRFILE...
            sCurrFile = Dir
        End If
    Loop
    RmDir sFolder

End Sub

Public Function GetFileOI(ByVal strFilename As String) As Boolean

  Dim Counter As Long

    If GetSetting(AV.AVname, "Settings", "Scan for Exec", True) Then
        For Counter = 1 To Len(FileTypesofInterrest) Step 3
            If InStr(1, strFilename, Mid$(FileTypesofInterrest, Counter, 3), vbTextCompare) <> 0 Then
                GetFileOI = True
                Exit Function
            End If
        Next Counter
    End If
    GetFileOI = False

End Function

Public Function IsFileaScript(strFilename As String) As Boolean

    IsFileaScript = False
    strFilename = UCase$(strFilename)
    If Mid$(strFilename, Len(strFilename) - 3, 3) = ".JS" Then
        IsFileaScript = True
     ElseIf Mid$(strFilename, Len(strFilename) - 3, 3) = "VBS" Then 'NOT MID$(STRFILENAME,...
        IsFileaScript = True
    End If
    Debug.Print Mid$(strFilename, Len(strFilename) - 3, 4)
End Function

Public Sub Log(ByVal strLog As String, typ As Integer, Optional ForceLog As Boolean)

    If GetSetting(AV.AVname, "Settings", "LogFile", "OFF") = "OFF" And ForceLog <> True Then
        Exit Sub
    End If
    Call frmLog.AddLog(typ, 0, strLog)
End Sub

Public Function UnzipFile2RandomPath(ByVal strFilename As String) As String

  Dim strPath As String
  Dim Counter As Long
  Dim Unzip   As cUnzip

    Randomize Timer
    strPath = "c:\temp" & Int(Rnd * 100) & "\"
    Set Unzip = New cUnzip
    With Unzip
        .ZipFile = strFilename
        .Directory
        For Counter = 1 To .FileCount
            .FileSelected(Counter) = True
        Next '  COUNTER '  COUNTER
    End With 'UNZIP
    Unzip.UnzipFolder = strPath
    Unzip.Unzip
    UnzipFile2RandomPath = strPath
    Set Unzip = Nothing
    DoEvents

End Function


