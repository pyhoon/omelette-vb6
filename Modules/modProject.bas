Attribute VB_Name = "modProject"
' Version : 0.x
' Author: Poon Yip Hoon
' Modified On : dd/MM/yyyy
' Descriptions : Project functions
' Dependencies:
' modTextFile
Option Explicit
Private i As Integer
'Private strProjectType As String
'Private strProjectFolder As String
'Private strProjectName As String
Private strProjectFormName() As String
Private strProjectFormFile() As String

Public Function ProjectExist(strProjectName As String, Optional strProjectPath As String) As Boolean
On Error GoTo Catch
    If Trim(strProjectName) = "" Then
        strProjectName = "Project1"
    End If
    If strProjectPath = "" Then strProjectPath = gstrProjectPath
    strProjectPath = gstrProjectPath & "\" & strProjectName
    If Dir(strProjectPath, vbDirectory) <> "" Then
        ProjectExist = True
    Else
        ProjectExist = False
    End If
    Exit Function
Catch:
    LogError "Error", "ProjectExist", Err.Description
End Function

Public Function CreateProject(strProjectName As String, strProjectType As String, Optional strProjectPath As String) As Boolean
On Error GoTo Catch
    If Trim(strProjectName) = "" Then
        strProjectName = "Project1"
    End If
    If strProjectPath = "" Then strProjectPath = gstrProjectPath
    strProjectPath = strProjectPath & "\" & strProjectName
    If Dir(strProjectPath, vbDirectory) <> "" Then
    
    Else
        MkDir strProjectPath
    End If
    If FileExists(strProjectPath & "\" & strProjectName & ".vbp") Then
    
    Else
        If strProjectType = "DEFAULT" Then
            WriteVbp strProjectPath & "\" & strProjectName & ".vbp", strProjectName
        ElseIf strProjectType = "STANDARD" Then
            ReDim strProjectFormName(0)
            strProjectFormName(0) = "frmForm1"
            ReDim strProjectFormFile(0)
            strProjectFormFile(0) = "frmForm1.frm"
            WriteFrm strProjectPath & "\" & strProjectFormFile(0), strProjectFormName(0), "Form1"
            WriteVbp strProjectPath & "\" & strProjectName & ".vbp", strProjectName, strProjectFormName(0), strProjectType
        Else ' BLANK
            WriteVbp strProjectPath & "\" & strProjectName & ".vbp", strProjectName
        End If
        'MsgBox "Project created successfully", vbInformation, "Successful"
        'gstrProjectName = strProjectName
        'gstrProjectPath = strProjectPath
        
        With mdiMain
            .mnuFileMakeExe.Caption = "&Compile " & gstrProjectName & ".exe..."
            .mnuFileMakeExeAndRun.Caption = "Compile and &Run " & gstrProjectName & ".exe..."
            .mnuFileMakeExe.Enabled = True
            .mnuFileMakeExeAndRun.Enabled = True
        End With
        'Unload Me
        With frmWindowProject
            .Width = 4000
            .Height = 4000
            .Left = mdiMain.Width - .Width - 300
            .Top = 0
            .Show
            '.RefreshTreeview strProjectType
        End With
    End If
    CreateProject = True
    Exit Function
Catch:
    LogError "Error", "CreateProject", Err.Description
    CreateProject = False
End Function

Public Sub WriteVbp(pstrFileName As String, _
                    Optional pstrProjectName As String, _
                    Optional pstrStartup As String, _
                    Optional pstrType As String)
    Dim FF As Integer
    Dim strText(28) As String
On Error GoTo Catch
    FF = FreeFile
    strText(0) = "Type=Exe"
    strText(1) = "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
    strText(1) = strText(1) & "C:\Windows\SysWOW64" ' System Directory
    strText(1) = strText(1) & "\stdole2.tlb"
    strText(1) = strText(1) & "#OLE Automation"
    If pstrType = "STANDARD" Then
        If pstrStartup = "" Then
            pstrStartup = "frmForm1" ' "Sub Main"
        End If
    Else
        If pstrStartup = "" Then
            pstrStartup = "Form1" ' "Sub Main"
        End If
    End If
    strText(2) = "Form=" & pstrStartup & ".frm"
    strText(3) = "Startup=""" & pstrStartup & """"
    strText(4) = "HelpFile="""""
    strText(5) = "Command32="""""
    If pstrProjectName = "" Then
        pstrProjectName = "Project1"
    End If
    strText(6) = "Name=""" & pstrProjectName & """"
    strText(7) = "HelpContextID=""0"""
    strText(8) = "CompatibleMode=""0"""
    strText(9) = "MajorVer=1"
    strText(10) = "MinorVer=0"
    strText(11) = "RevisionVer=0"
    strText(12) = "AutoIncrementVer=0"
    strText(13) = "ServerSupportFiles=0"
    strText(14) = "VersionCompanyName=""" & App.CompanyName & """" ' "Computerise System Solutions"
    strText(15) = "CompilationType=0"
    strText(16) = "OptimizationType=0"
    strText(17) = "FavorPentiumPro(tm)=0"
    strText(18) = "CodeViewDebugInfo=0"
    strText(19) = "NoAliasing=0"
    strText(20) = "BoundsCheck=0"
    strText(21) = "OverflowCheck=0"
    strText(22) = "FlPointCheck=0"
    strText(22) = "FDIVCheck=0"
    strText(23) = "UnroundedFP=0"
    strText(24) = "StartMode=0"
    strText(25) = "Unattended=0"
    strText(26) = "Retained=0"
    strText(27) = "ThreadPerObject=0"
    strText(28) = "MaxNumberOfThreads=1"
    Open pstrFileName For Output As #FF
        For i = 0 To 28
            Print #FF, strText(i)
        Next
    Close #FF
    Exit Sub
Catch:
    LogError "Error", "WriteVbp(" & pstrFileName & ")", Err.Description
End Sub

' Write the file (overwrite existing file)
Public Sub WriteFrm(pstrFileName As String, _
                    Optional pstrVB_Name As String, _
                    Optional pstrCaption As String)
    Dim FF As Integer
    Dim strText(16) As String
On Error GoTo Catch
    FF = FreeFile
    strText(0) = "Version 5.00"
    If pstrVB_Name = "" Then
        pstrVB_Name = "frmForm1"
    End If
    strText(1) = "Begin VB.Form " & pstrVB_Name
    If pstrCaption = "" Then
        pstrCaption = "Form1"
    End If
    strText(2) = "   Caption         =   """ & pstrCaption & """"
    strText(3) = "   ClientHeight    =   4275"
    strText(4) = "   ClientLeft      =   14985"
    strText(5) = "   ClientTop       =   1455"
    strText(6) = "   ClientWidth     =   7260"
    strText(7) = "   LinkTopic       =   ""Form1"""
    strText(8) = "   ScaleHeight     =   4275"
    strText(9) = "   ScaleWidth      =   7260"
    strText(10) = "   StartUpPosition = 1    'CenterOwner"
    strText(11) = "End"
    If pstrVB_Name = "" Then
        pstrVB_Name = pstrFileName
    End If
    strText(12) = "Attribute VB_Name = """ & pstrVB_Name & """"
    strText(13) = "Attribute VB_GlobalNameSpace = False"
    strText(14) = "Attribute VB_Creatable = False"
    strText(15) = "Attribute VB_PredeclaredId = True"
    strText(16) = "Attribute VB_Exposed = False"
    Open pstrFileName For Output As #FF
        For i = 0 To 16
            Print #FF, strText(i)
        Next
    Close #FF
    Exit Sub
Catch:
    LogError "Error", "WriteFrm(" & pstrFileName & ")", Err.Description
End Sub

Public Sub WriteBas(pstrFileName As String, _
                    Optional pstrVB_Name As String)
    Dim FF As Integer
    Dim strText(1) As String
On Error GoTo Catch
    FF = FreeFile
    strText(0) = "Attribute VB_Name = """ & pstrVB_Name & """"
    strText(1) = "Option Explicit"
    Open pstrFileName For Output As #FF
        For i = 0 To 1
            Print #FF, strText(i)
        Next
    Close #FF
    Exit Sub
Catch:
    LogError "Error", "WriteBas(" & pstrFileName & ")", Err.Description
End Sub

' Continue write existing file - CRUD functions
Public Sub WriteBas0001(pstrFileName As String)
    Dim FF As Integer
    Dim strText(6) As String
On Error GoTo Catch
    FF = FreeFile
    strText(0) = "Public Function AddNew() As Boolean"
    strText(1) = "    AddNew = True"
    strText(2) = "End Function"
    strText(3) = ""
    strText(4) = "Public Function DeleteID() As Boolean"
    strText(5) = "    DeleteID = True"
    strText(6) = "End Function"
    Open pstrFileName For Output As #FF
        For i = 0 To 6
            Print #FF, strText(i)
        Next
    Close #FF
    Exit Sub
Catch:
    LogError "Error", "WriteBas0001(" & pstrFileName & ")", Err.Description
End Sub

