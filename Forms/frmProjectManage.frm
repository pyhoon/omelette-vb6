VERSION 5.00
Begin VB.Form frmProjectManage 
   Caption         =   "Manage Project"
   ClientHeight    =   2415
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   7575
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtProjectData 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      MaxLength       =   255
      TabIndex        =   7
      Top             =   960
      Width           =   5535
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtProjectFile 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      MaxLength       =   255
      TabIndex        =   9
      Top             =   1320
      Width           =   5535
   End
   Begin VB.TextBox txtProjectPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      MaxLength       =   255
      TabIndex        =   5
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox txtProjectName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1800
      MaxLength       =   255
      TabIndex        =   3
      Top             =   240
      Width           =   5535
   End
   Begin VB.Label Label3 
      Caption         =   "Project Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   1500
   End
   Begin VB.Label Label4 
      Caption         =   "Project File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "Project Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "Project Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1500
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
Try:
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
            .CloseRs rst
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
        .CloseRs rst
        .CloseMdb
    End With
    Exit Sub
Catch:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "PopulateData"
    LogError "Error", "PopulateData", Err.Description
    With DB
        .CloseRs rst
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
        Set rst = .OpenRs(gstrSQL)
        If .ErrorDesc <> "" Then
            MsgBox "Error: " & .ErrorDesc, vbExclamation, "SaveData"
            .CloseRs rst
            .CloseMdb
            Exit Sub
        End If
        If rst.EOF Then
            MsgBox "Project not found!", vbExclamation, "SaveData"
            .CloseRs rst
            .CloseMdb
            Exit Sub
        End If
        .CloseRs rst
        'Update Project table
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
        .CloseRs rst
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
