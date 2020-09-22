Attribute VB_Name = "modError"
Public Sub ErrorFunc(Err_Number As Integer, Err_Description As String, Err_Routine As String, Optional RoutineVariables As String)
Debug.Print Now & " Error occured! System halted"
Log LoadResString(154) & ": " & Err_Description & " " & Err_Routine, 3, True
MsgBox LoadResString(153) & ": " & Err_Description, vbCritical + vbOKOnly
End Sub
