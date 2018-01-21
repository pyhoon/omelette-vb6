VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Omelette"
   ClientHeight    =   7995
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13845
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7740
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23892
            Object.ToolTipText     =   "Project Name"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewProject 
         Caption         =   "&New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenProject 
         Caption         =   "&Open Project..."
         Shortcut        =   ^O
      End
      Begin VB.Menu Separator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMakeExe 
         Caption         =   "&Compile App.exe..."
         Enabled         =   0   'False
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuFileMakeExeAndRun 
         Caption         =   "Compile and &Run App.exe..."
         Enabled         =   0   'False
         Shortcut        =   ^{F8}
      End
      Begin VB.Menu Separator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewProjectExplorer 
         Caption         =   "&Project Explorer"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
      Begin VB.Menu mnuFileManageProject 
         Caption         =   "&Manage Project"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuRunStart 
         Caption         =   "&Start"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "Q&uery"
   End
   Begin VB.Menu mnuDiagram 
      Caption         =   "D&iagram"
   End
   Begin VB.Menu mnuAddIns 
      Caption         =   "&Add-Ins"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu Separator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCommandLine 
         Caption         =   "&Command Line..."
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Omelette..."
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
    
    Private Const SYNCHRONIZE = &H100000
    Private Const INFINITE = -1&

    ' Const COMPILER = "C:\Program Files\Microsoft Visual Studio\Vb98\Vb6.exe"
    Private Const COMPILER = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.exe"
    'Const PROJECT_NAME = "Target"
    
    Private strProjectPath As String
    Private strProjectVbp As String
    Private strProjectExe As String
    Private cmd As String

' Check function mnuRunStart_Click()

Private Sub MDIForm_Load()
    'strProjectPath = App.Path & "\Projects\"
    InitGlobalVariables
    Me.WindowState = vbMaximized
End Sub

Private Sub mnuAbout_Click()
    MsgBox App.ProductName & vbCrLf & _
    "Version " & App.Major & "." & App.Minor & " Build " & App.Revision & vbCrLf & vbCrLf & _
    "Copyright" & Chr(169) & " 2017-2018 " & App.CompanyName & "." & vbCrLf & _
    "All rights reserved.", vbInformation, App.Title
End Sub

Private Sub mnuCommandLine_Click()
    With frmCommandLine
        .Show 'vbModal
    End With
End Sub

Private Sub mnuExit_Click()
    ' Show close dialog
    End
End Sub

Private Sub mnuFileMakeExe_Click()
    CompileExe
End Sub

Private Sub mnuFileMakeExeAndRun_Click()
    CompileExe
    RunExe
End Sub

Private Sub mnuFileManageProject_Click()
    With frmProjectManage
        .Show vbModal
    End With
End Sub

Private Sub mnuFileNewProject_Click()
    With frmProjectNew
        .Show vbModal
    End With
End Sub

Private Sub mnuFileOpenProject_Click()
    With frmProjectOpen
        .Show vbModal
    End With
End Sub

Private Sub mnuOptions_Click()
    'frmOptions.Show
End Sub

Private Sub mnuRun_Click()
    If strProjectExe = "" Then
        mnuRunStart.Enabled = False
    End If
End Sub

Private Sub mnuRunStart_Click()
    If strProjectExe = "" Then
        'With frmLogin ' frmLiveApp
        '    .Show
        '    .ZOrder 0
        'End With
        MsgBox "Executable file not set!", vbExclamation, "Error"
    Else
        'strProjectExe = App.Path & "\Projects\Test\Test.exe"
        RunExe
    End If
End Sub

Private Sub mnuViewProjectExplorer_Click()
    With frmWindowProject
        .Width = 4000
        .Height = 4000
        .Left = Me.Width - .Width - 300
        .Top = 0
        .Show
    End With
End Sub

' Initialize global variables
Private Sub InitGlobalVariables()
    gstrProjectName = ""
    'gstrProjectFile = ""
    
    gstrProjectFolder = "\Projects\" ' default
    gstrProjectPath = App.Path & gstrProjectFolder
    gstrProjectData = "Data\"
    gstrProjectDataPath = "" ' gstrProjectPath & gstrProjectData
    gstrProjectDataFile = "Data.mdb"
    gstrProjectDataPassword = ""
    gstrProjectItemsFile = "Items.mdb"
    gstrProjectItemsPassword = ""
    
    gstrMasterFolder = "\"
    gstrMasterPath = App.Path & gstrMasterFolder
    gstrMasterData = "Databases\"
    gstrMasterDataPath = gstrMasterPath & gstrMasterData
    gstrMasterDataFile = "Projects.mdb"
    gstrMasterDataPassword = ""
End Sub

' Source: http://www.vb-helper.com/howto_make_and_run.html
' Start the indicated program and wait for it
' to finish, hiding while we wait.
Private Sub ShellAndWait(ByVal program_name As String)
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

Private Sub CompileExe()
    If gstrProjectPath = "" Then Exit Sub
    strProjectPath = gstrProjectPath
    ' Get the project file name.
    If Right$(strProjectPath, 1) <> "\" Then strProjectPath = strProjectPath & "\"
    strProjectVbp = strProjectPath & gstrProjectName & ".vbp"
    strProjectExe = strProjectPath & gstrProjectName & ".exe"

    ' Compose the compile command.
    cmd = """" & COMPILER & """ /MAKE """ & strProjectVbp & """"

    ' Shell this command and wait for it to finish.
    ShellAndWait cmd
End Sub

Private Sub RunExe()
    If strProjectExe = "" Then Exit Sub
    ' Execute the newly compiled program.
    cmd = """" & strProjectExe & """"
    'ShellAndWait cmd
    Shell cmd, vbNormalFocus
End Sub

'Private Sub mnuViewProjectUserAdmin_Click()
'    With frmUserAdmin
'        .Show
'        .ZOrder 0
'    End With
'End Sub
