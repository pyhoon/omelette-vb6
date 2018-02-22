VERSION 5.00
Begin VB.Form frmProjectOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Project"
   ClientHeight    =   4815
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Project Properties"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   5775
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Project Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.ListBox lstProject 
      Height          =   450
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   120
      Pattern         =   "*.vbp"
      TabIndex        =   5
      Top             =   2520
      Width           =   5775
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5775
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Drive"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmProjectOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strProjectName As String

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOpen_Click()
    If Not File1.FileName Like "*.vbp" Then ' strProjectName = "" Then
        MsgBox "Please select a project", vbInformation, "Open Project"
        Exit Sub
    End If
    If SearchProject(Dir1.Path, File1.FileName, strProjectName) Then
        If strProjectName <> "" Then
            OpenProject Dir1.Path, File1.FileName, strProjectName
            Unload Me
        End If
    End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
    If File1.FileName Like "*.vbp" Then
        ' Search this project and path in database
        If SearchProject(Dir1.Path, File1.FileName, strProjectName) Then
            Label2.Caption = "Project Name: " & strProjectName
        Else
            Label2.Caption = "Project Not found"
        End If
    End If
End Sub

Private Sub Form_Load()
    Dir1.Path = App.Path & "\Projects\"
End Sub

Private Function SearchProject(ByVal strProjectPath As String, ByVal strProjectFile As String, ByRef strProjectName As String) As Boolean
    Dim DB As New OmlDatabase
    Dim rst As ADODB.Recordset
    With DB
        .DataPath = gstrMasterDataPath & "\"
        .DataFile = gstrMasterDataFile
        '.DataPath = gstrProjectDataPath & "\"
        '.DataFile = gstrProjectItemsFile
        .OpenMdb
        'SQL_SELECT
        'SQL_FROM "Project"
        SQL_SELECT_ALL "Project"
        SQL_WHERE_Text "ProjectPath", strProjectPath
        SQL_AND_Text "ProjectFile", strProjectFile
        Set rst = .OpenRs(gstrSQL)
        If .ErrorDesc <> "" Then
            LogError "Error", "SearchProject/frmProjectOpen", .ErrorDesc
            .CloseRS rst
            .CloseMdb
            SearchProject = False
            Exit Function
        End If
        If rst Is Nothing Or rst.EOF Then
            .CloseRS rst
            .CloseMdb
            SearchProject = False
            Exit Function
        End If
        With rst
            If Not .EOF Then
                strProjectName = !ProjectName
                SearchProject = True
            Else
                SearchProject = False
            End If
        End With
        .CloseRS rst
        .CloseMdb
    End With
End Function

Private Sub OpenProject(ByVal strProjectPath As String, ByVal strProjectFile As String, ByVal strProjectName As String)
    gstrProjectName = strProjectName
    gstrProjectPath = strProjectPath
    gstrProjectFile = strProjectFile
    gstrProjectDataPath = strProjectPath & "\" & gstrProjectData '?
    'gstrProjectDataPath = gstrProjectPath & gstrProjectName & "\" & gstrProjectData
    With mdiMain
        .Caption = App.Title & " - [" & strProjectName & "]"
        .sbStatus.Panels(1).Text = strProjectPath 'Dir1.Path
        .mnuFileManageProject.Enabled = True
        .mnuFileMakeExe.Caption = "&Compile " & strProjectName & ".exe..."
        .mnuFileMakeExeAndRun.Caption = "Compile and &Run " & strProjectName & ".exe..."
        .mnuFileMakeExe.Enabled = True
        .mnuFileMakeExeAndRun.Enabled = True
    End With
    Unload Me
    Unload frmWindowProject
    With frmWindowProject
        .Width = 4000
        .Height = 4000
        .Left = mdiMain.Width - .Width - 300
        .Top = 0
        .Show
        '.RefreshTreeview
    End With
End Sub
