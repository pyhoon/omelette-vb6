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
      Left            =   4680
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
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
    If Not File1.FileName Like "*.vbp" Then
        MsgBox "Please select a project", vbInformation, "Open Project"
        Exit Sub
    End If
    If FileExists(File1.Path & "\" & File1.FileName) Then
        OpenProject strProjectName, File1.Path, File1.FileName
        Unload Me
    End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Dir1_Click()
    File1.Path = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub File1_Click()
    If File1.FileName Like "*.vbp" Then
        ' Try get project name from vbp file
        SearchTextFile File1.Path & "\" & File1.FileName, "Name=", strProjectName
        If strProjectName <> "" Then
            strProjectName = Replace(strProjectName, "Name=", "")   ' Remove Name=
            strProjectName = Replace(strProjectName, """", "")      ' Remove double quotes
            Label2.Caption = "Project Name: " & strProjectName
        Else
            Label2.Caption = "" '"Project Not found"
        End If
    End If
End Sub

Private Sub Form_Load()
    Dir1.Path = App.Path & "\Projects\"
End Sub

Private Sub OpenProject(ByVal strProjectName As String, _
                        ByVal strProjectPath As String, _
                        ByVal strProjectFile As String)
    gstrProjectName = strProjectName
    gstrProjectPath = strProjectPath
    gstrProjectFile = strProjectFile
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
