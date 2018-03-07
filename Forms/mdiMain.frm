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
         Enabled         =   0   'False
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
    Dim strProjectExe As String
    
    strProjectExe = App.Path & "\" & gstrProjectFolder & "\" & gstrProjectName & "\" & gstrProjectName & ".exe"
    If FileExists(strProjectExe) Then
        mnuRunStart.Enabled = True
    Else
        mnuRunStart.Enabled = False
    End If
End Sub

Private Sub mnuRunStart_Click()
    RunExe
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
    
    gstrProjectFolder = "Projects"
    gstrProjectPath = App.Path & "\" & gstrProjectFolder
    gstrProjectData = "Storage"
    gstrProjectDataPath = "" ' gstrProjectPath & "\" & gstrProjectName & "\" & gstrProjectData
    gstrProjectDataFile = "Data.mdb"
    gstrProjectDataPassword = ""
    gstrProjectItemsFile = "Items.mdb"
    gstrProjectItemsPassword = ""
    
    'gstrMasterFolder = "\"
    gstrMasterPath = App.Path '& gstrMasterFolder
    'gstrMasterFolder = gstrMasterPath & "\" & gstrMasterFolder
    gstrMasterData = "Databases"
    gstrMasterDataPath = gstrMasterPath & "\" & gstrMasterData
    gstrMasterDataFile = "Projects.mdb"
    gstrMasterDataPassword = ""
End Sub

'Private Sub mnuViewProjectUserAdmin_Click()
'    With frmUserAdmin
'        .Show
'        .ZOrder 0
'    End With
'End Sub
