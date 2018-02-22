VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProjectManage 
   Caption         =   "Manage Project"
   ClientHeight    =   1710
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   8730
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtProjectData 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   255
      TabIndex        =   6
      Top             =   840
      Width           =   5535
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtProjectFile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   255
      TabIndex        =   8
      Top             =   1200
      Width           =   5535
   End
   Begin VB.TextBox txtProjectPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   255
      TabIndex        =   4
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox txtProjectName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1560
      MaxLength       =   255
      TabIndex        =   2
      Top             =   120
      Width           =   5535
   End
   Begin MSComctlLib.ListView lvwProjects 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Project Name"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Project Path"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Project File"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Project Data"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Project File"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Project Path"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Project Name"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Visible         =   0   'False
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
   End
End
Attribute VB_Name = "frmProjectManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    PopulateData
End Sub

Private Sub cmdSave_Click()
    SaveData
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub PopulateData()
Dim DB As New OmlDatabase
Dim rst As ADODB.Recordset
    If gstrProjectDataPath = "" Or gstrProjectItemsFile = "" Then Exit Sub
On Error GoTo Catch
    SQL_SELECT
    SQLText "ProjectName"
    SQLText "ProjectPath"
    SQLText "ProjectData"
    SQLText "ProjectFile"
    SQLText "CreatedDate"
    SQLText "CreatedBy", False
    SQL_FROM "Project"
    With DB
        .DataPath = gstrProjectDataPath & "\"
        .DataFile = gstrProjectItemsFile
        .OpenMdb
        Set rst = .OpenRs(gstrSQL)
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "PopulateData"
            .CloseRS rst
            .CloseMdb
            Exit Sub
        End If
        If Not rst.EOF Then
            SetTextBoxValue txtProjectName, rst!ProjectName
            SetTextBoxValue txtProjectPath, rst!ProjectPath
            SetTextBoxValue txtProjectData, rst!ProjectData
            SetTextBoxValue txtProjectFile, rst!ProjectFile
        Else
            SetTextBoxValue txtProjectName
            SetTextBoxValue txtProjectPath
            SetTextBoxValue txtProjectData
            SetTextBoxValue txtProjectFile
        End If
        .CloseRS rst
        .CloseMdb
    End With
    Exit Sub
Catch:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "PopulateData"
    LogError "Error", "PopulateData", Err.Description
    With DB
        .CloseRS rst
        .CloseMdb
    End With
End Sub

Private Sub SaveData()
Dim DB As New OmlDatabase
Dim rst As ADODB.Recordset
On Error GoTo Catch
    If vbNo = MsgBox("Do you want to Save this Project?", vbQuestion + vbYesNo, "SaveData") Then
        Exit Sub
    End If
    If Trim(txtProjectName.Text) = "" Then
        MsgBox "Please key in Project Name", vbExclamation, "SaveData"
        txtProjectName.SetFocus
        Exit Sub
    End If
    If Trim(txtProjectPath.Text) = "" Then
        MsgBox "Please key in Project Path", vbExclamation, "SaveData"
        txtProjectPath.SetFocus
        Exit Sub
    End If
    If Trim(txtProjectData.Text) = "" Then
        MsgBox "Please key in Project Data", vbExclamation, "SaveData"
        txtProjectData.SetFocus
        Exit Sub
    End If
    If Trim(txtProjectFile.Text) = "" Then
        MsgBox "Please key in Project File", vbExclamation, "SaveData"
        txtProjectFile.SetFocus
        Exit Sub
    End If
    With DB
        .DataPath = gstrProjectDataPath & "\"
        .DataFile = gstrProjectItemsFile
        .OpenMdb
        SQL_SELECT_ALL "Project"
        SQL_WHERE_Text "ProjectName", gstrProjectName
        'SQL_AND_Text "ProjectPath", gstrProjectPath & "\" & gstrProjectName
        Set rst = .OpenRs(gstrSQL)
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "SaveData"
            .CloseRS rst
            .CloseMdb
            Exit Sub
        End If
        If rst.EOF Then
            MsgBox "Project not found!", vbExclamation, "SaveData"
            .CloseRS rst
            .CloseMdb
            Exit Sub
        End If
        .CloseRS rst
        'Update table Project
        SQL_UPDATE "Project"
        SQL_SET_Text "ProjectName", Trim(txtProjectName.Text)
        SQL_SET_Text "ProjectPath", Trim(txtProjectPath.Text)
        SQL_SET_Text "ProjectData", Trim(txtProjectData.Text)
        SQL_SET_Text "ProjectFile", Trim(txtProjectFile.Text)
        SQL_SET_DateTime "ModifiedDate", Now
        SQL_SET_Text "ModifiedBy", "Aeric", False
        .Execute gstrSQL ', lngRecordsAffected
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "SaveData"
            .CloseMdb
            Exit Sub
        End If
        .CloseMdb
        MsgBox "Project is saved!", vbInformation, "SaveData"
    End With
    Exit Sub
Catch:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "SaveData"
    LogError "Error", "SaveData", Err.Description
    With DB
        .CloseRS rst
        .CloseMdb
    End With
End Sub

Private Sub SetTextBoxValue(ByVal ctrlTextbox As TextBox, Optional ByVal objData As Object)
    If objData Is Nothing Then
        ctrlTextbox.Text = ""
        Exit Sub
    End If
    If objData <> "" Then
        ctrlTextbox.Text = objData
    Else
        ctrlTextbox.Text = ""
    End If
End Sub
