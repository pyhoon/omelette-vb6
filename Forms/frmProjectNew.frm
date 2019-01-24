VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProjectNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Project"
   ClientHeight    =   3195
   ClientLeft      =   7080
   ClientTop       =   4215
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList istIcons 
      Left            =   3720
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectNew.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectNew.frx":10DA
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectNew.frx":21B4
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectNew.frx":328E
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectNew.frx":4368
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectNew.frx":5442
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectNew.frx":651C
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProjectNew.frx":75F6
            Key             =   "IMG8"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwProjects 
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5106
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      Icons           =   "istIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmProjectNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strNewProjectName As String
Private strProjectType As String
    
Private Sub Form_Load()
    ListTemplatesFromDatabase
End Sub

Private Sub cmdOK_Click()
    strNewProjectName = InputBox("Please enter your project name:", "New project", "Project1")
    If Trim(strNewProjectName) = "" Then
        MsgBox "Invalid Project Name!", vbExclamation, "New Project"
        Exit Sub
    End If
    'Unload Me
    If ProjectExist(strNewProjectName, "") Then
        MsgBox "Project already exist!", vbExclamation, "New Project"
        Exit Sub
    End If
    CreateProject strProjectType, "", strNewProjectName
    gstrProjectName = strNewProjectName
    ' Open Project
    OpenProject strNewProjectName, gstrProjectPath
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    MsgBox "No help file", vbInformation, "Help"
End Sub

Private Sub lvwProjects_DblClick()
    strNewProjectName = InputBox("Please enter your project name:", "New project", "Project1")
    If Trim(strNewProjectName) = "" Then
        MsgBox "Invalid Project Name!", vbExclamation, "New Project"
        Exit Sub
    End If
    If ProjectExist(strNewProjectName, "") Then
        MsgBox "Project already exist!", vbExclamation, "New Project"
        Exit Sub
    End If
    CreateProject strProjectType, "", strNewProjectName
    gstrProjectName = strNewProjectName
    ' Open Project
    OpenProject strNewProjectName, gstrProjectPath
End Sub

Private Sub lvwProjects_ItemClick(ByVal Item As MSComctlLib.ListItem)
    strProjectType = lvwProjects.SelectedItem.Key
End Sub

Private Sub ListTemplatesFromDatabase()
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
Dim rst As ADODB.Recordset
Dim i As Integer
On Error GoTo Catch
Try:
    DB.DataPath = gstrMasterDataPath & "\"
    DB.DataFile = gstrMasterDataFile
    DB.OpenMdb
    SQ.SELECT_ALL "Template"
    SQ.WHERE_Boolean "Active", True
    Set rst = DB.OpenRs(SQ.Text)
    If DB.ErrorDesc <> "" Then
        LogError "Error", "ListTemplatesFromDatabase/frmProjectNew", DB.ErrorDesc
        DB.CloseRs rst
        DB.CloseMdb
        Exit Sub
    End If
    If Not (rst Is Nothing Or rst.EOF) Then
        While Not rst.EOF
            i = i + 1
            lvwProjects.ListItems.Add i, rst!TemplateKey, rst!TemplateText, CInt(rst!TemplateIcon)
            rst.MoveNext
        Wend
    End If
    DB.CloseRs rst
    DB.CloseMdb
    Exit Sub
Catch:
    LogError "Error", "ListTemplatesFromDatabase/frmProjectNew->", Err.Description
    DB.CloseRs rst
    DB.CloseMdb
End Sub

Private Sub OpenProject(ByVal strProjectName As String, ByVal strProjectPath As String)
    gstrProjectName = strProjectName
    gstrProjectPath = strProjectPath
    gstrProjectFile = strProjectName & ".vbp"
    gstrProjectDataPath = gstrProjectPath & "\" & gstrProjectData
    With mdiMain
        .Caption = App.Title & " - [" & strProjectName & "]"
        .sbStatus.Panels(1).Text = gstrProjectPath
        .mnuFileManageProject.Enabled = True
        .mnuFileMakeExe.Caption = "&Compile " & strProjectName & ".exe..."
        .mnuFileMakeExeAndRun.Caption = "Compile and &Run " & strProjectName & ".exe..."
        .mnuFileMakeExe.Enabled = True
        .mnuFileMakeExeAndRun.Enabled = True
        With .tbrMain
            .Buttons(3).Enabled = True
            .Buttons(4).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(6).Enabled = True
        End With
    End With
    Unload Me
    'Unload frmWindowProject
    With frmWindowProject
        .Width = 4000
        .Height = 4000
        .Left = mdiMain.Width - .Width - 300
        .Top = 0
        .Show
        .ReadItemsData
    End With
End Sub
