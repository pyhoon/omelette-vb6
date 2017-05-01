Attribute VB_Name = "modTextFile"
Option Explicit

Public Sub LogError(pstrFileName As String, pstrNote As String, Optional pstrError As String)
    Dim SEPARATOR As Variant
    Dim NEWLINE As Variant
    
    SEPARATOR = vbTab ' vbCrLf
    NEWLINE = vbCrLf
On Error GoTo Catch
    Open pstrFileName For Append As #1
    If pstrError = "" Then
        Print #1, NEWLINE & FormatDateAndTime(Now) & SEPARATOR & pstrNote
    Else
        Print #1, NEWLINE & FormatDateAndTime(Now) & SEPARATOR & pstrNote & SEPARATOR & pstrError
    End If
    Close #1
    Exit Sub
Catch:
    MsgBox "Error #" & Err.Number & NEWLINE & Err.Description, vbExclamation, App.Title
End Sub

Public Function WriteTextFile(pstrFileName As String, pstrText As String)
    Dim FF As Integer
On Error GoTo Catch
    FF = FreeFile
    Open pstrFileName For Output As #FF
        Print #FF, CStr(pstrText)
    Close #FF
    Exit Function
Catch:
    LogError "Error", "WriteTextFile(" & pstrFileName & ")", Err.Description
End Function

Public Sub ReadTextFile(ByVal pstrFileName As String, ByVal pintLineNo As Integer, ByRef pstrOutput As String)
Dim FF As Integer
Dim i As Integer
On Error GoTo RE
    FF = FreeFile
    Open pstrFileName For Input As #FF
        If pintLineNo < 0 Then
            Do Until EOF(FF) = True
                Input #FF, pstrOutput
            Loop
        ElseIf pintLineNo > 0 Then
            For i = 0 To pintLineNo
                If Not EOF(FF) Then
                    Input #FF, pstrOutput
                Else
                    Exit For
                End If
            Next
        Else
            Input #FF, pstrOutput
        End If
    Close
    Exit Sub
RE:
    LogError "Error", "ReadTextFile(" & pstrFileName & ", " & pintLineNo & ")", Err.Description
End Sub
