Attribute VB_Name = "modProject"
' Version : 0.3
' Author: Poon Yip Hoon
' Modified On : 07/03/2018
' Modified On : 29/06/2018
' Descriptions : Project functions
' Dependencies:
' modTextFile
Option Explicit
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    
    Private Const SYNCHRONIZE = &H100000
    Private Const INFINITE = -1&
    Public Const COMPILER = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.exe"

    Private i As Integer
    Private strProjectPath As String
    Private strProjectVbp As String
    Private strProjectExe As String
    Private cmd As String

Public Function ProjectExist(strProjectName As String, _
                            Optional strProjectPath As String) As Boolean
On Error GoTo Catch
    If Trim(strProjectName) = "" Then
        strProjectName = "Project1"
    End If
    If strProjectPath = "" Then strProjectPath = gstrMasterPath & "\" & gstrProjectFolder & "\" & strProjectName
    If Dir(strProjectPath, vbDirectory) <> "" Then
        ProjectExist = True
    Else
        ProjectExist = False
    End If
    Exit Function
Catch:
    LogError "Error", "ProjectExist", Err.Description
End Function

' =====================================
' Moved from mdiMain and set as Public
' =====================================
' Source: http://www.vb-helper.com/howto_make_and_run.html
' Start the indicated program and wait for it
' to finish, hiding while we wait.
Public Sub ShellAndWait(ByVal program_name As String)
    Dim process_id As Long
    Dim process_handle As Long
' Start the program.
On Error GoTo ShellError
    process_id = Shell(program_name, vbNormalFocus)
    On Error GoTo 0

    ' Hide.
    'Me.Visible = False
    DoEvents

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    ' Reappear.
    'Me.Visible = True
    
    Exit Sub
ShellError:
    MsgBox "Error running '" & program_name & "'" & vbCrLf & Err.Description
End Sub

Public Sub CompileExe()
    If gstrProjectName = "" Or gstrProjectFile = "" Then Exit Sub
    strProjectPath = gstrMasterPath & "\" & gstrProjectFolder & "\" & gstrProjectName
    ' Get the project file name.
    'If Right$(strProjectPath, 1) <> "\" Then strProjectPath = strProjectPath & "\"
    strProjectVbp = strProjectPath & "\" & gstrProjectName & ".vbp"
    strProjectExe = strProjectPath & "\" & gstrProjectName & ".exe"
    
    If FileExists(strProjectVbp) = True Then
        ' Compose the compile command.
        cmd = """" & COMPILER & """ /MAKE """ & strProjectVbp & """"
    
        ' Shell this command and wait for it to finish.
        ShellAndWait cmd
    End If
End Sub

Public Sub RunExe()
    'If strProjectExe = "" Then Exit Sub
    strProjectPath = gstrMasterPath & "\" & gstrProjectFolder & "\" & gstrProjectName
    strProjectExe = strProjectPath & "\" & gstrProjectName & ".exe"
    If Not FileExists(strProjectExe) Then Exit Sub
    ' Execute the newly compiled program.
    cmd = """" & strProjectExe & """"
    'ShellAndWait cmd
    Shell cmd, vbNormalFocus
End Sub
