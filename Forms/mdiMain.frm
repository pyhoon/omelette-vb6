VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Omelette"
   ClientHeight    =   7995
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13845
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13845
      _ExtentX        =   24421
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "istIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Open project directory"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Open project in Visual Basic"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Explorer"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Properties"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Command"
            ImageIndex      =   7
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList istIcons 
         Left            =   12600
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":0542
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":0A84
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":0B96
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":0CA8
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":0DBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "mdiMain.frx":0ECC
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
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
      Visible         =   0   'False
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
      Begin VB.Menu mnuAddForm 
         Caption         =   "Add &Form"
      End
      Begin VB.Menu mnuAddModule 
         Caption         =   "Add &Module"
      End
      Begin VB.Menu mnuAddClass 
         Caption         =   "Add &Class"
      End
      Begin VB.Menu Separator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileManageProject 
         Caption         =   "Manage &Project"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Enabled         =   0   'False
      Visible         =   0   'False
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
      Visible         =   0   'False
   End
   Begin VB.Menu mnuDiagram 
      Caption         =   "D&iagram"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuAddIns 
      Caption         =   "&Add-Ins"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu mnuImportFileToTemplate 
         Caption         =   "&Import File to Template..."
      End
      Begin VB.Menu Separator4 
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
'Private Const COMPILER = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.exe"

Private Sub MDIForm_Load()
    'strProjectPath = App.Path & "\Projects\"
    InitGlobalVariables
    Me.WindowState = vbMaximized
End Sub

Private Sub mnuAbout_Click()
    MsgBox App.ProductName & vbCrLf & _
    "Version " & App.Major & "." & App.Minor & " Build " & App.Revision & vbCrLf & vbCrLf & _
    "Copyright" & Chr(169) & " 2017-2019 " & App.CompanyName & "." & vbCrLf & _
    "All rights reserved.", vbInformation, App.Title
End Sub

Private Sub mnuAddForm_Click()
    ' Add Form
    
    ' 1 Check if form exist
    
    ' Update Items
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

Private Sub mnuImportFileToTemplate_Click()
    With frmImportTemplate
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
    'gstrProjectFile = ""
    gstrProjectName = ""
    gstrProjectFolder = "Projects" ' use gstrMasterProject
    gstrProjectPath = "" ' App.Path & "\" & gstrProjectFolder
    gstrProjectData = "Storage"
    gstrProjectDataPath = ""
    gstrProjectDataFile = "Data.mdb"
    gstrProjectDataPassword = ""
    gstrProjectItemsFile = "Items.mdb"
    gstrProjectItemsPassword = ""
    'gstrProjectForms = "Forms"
    gstrProjectModules = "Modules"
    gstrProjectClasses = "Classes"

    gstrMasterPath = App.Path
    'gstrMasterProject = "Projects" ' should use this
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

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' NEW
            mnuFileNewProject_Click
        Case 2 ' OPEN
            mnuFileOpenProject_Click
        Case 3 ' PATH
            'mnuRunStart_Click
            If gstrProjectPath <> "" Then
                Shell "Explorer " & gstrProjectPath
            End If
        Case 4 ' PROJECT
            If gstrProjectPath <> "" Then
                Shell COMPILER & " " & gstrProjectPath & "\" & gstrProjectName & ".vbp", vbNormalFocus
            End If
        Case 5
            mnuViewProjectExplorer_Click
        Case 6
            mnuFileManageProject_Click
        Case 7
            mnuCommandLine_Click
    End Select
End Sub
