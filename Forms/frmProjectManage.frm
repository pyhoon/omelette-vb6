VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProjectManage 
   Caption         =   "Manage Project"
   ClientHeight    =   1290
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   8730
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtProjectFile 
      Height          =   285
      Left            =   1560
      MaxLength       =   255
      TabIndex        =   6
      Top             =   840
      Width           =   5535
   End
   Begin VB.TextBox txtProjectPath 
      Height          =   285
      Left            =   1560
      MaxLength       =   255
      TabIndex        =   5
      Top             =   480
      Width           =   5535
   End
   Begin VB.TextBox txtProjectName 
      Height          =   285
      Left            =   1560
      MaxLength       =   255
      TabIndex        =   4
      Top             =   120
      Width           =   5535
   End
   Begin MSComctlLib.ListView lvwProjects 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   1320
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
      Caption         =   "Project File"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Project Path"
      Height          =   255
      Left            =   240
      TabIndex        =   2
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
'Dim strSQL As String

Private Sub Form_Load()
    'ListProjects
    PopulateData
End Sub

'Private Sub lvwProjects_Click()
'    If lvwProjects.ListItems.Count = 0 Then
'        Exit Sub
'    End If
'    PopulateValues lvwProjects.SelectedItem.Text
'End Sub

'Private Sub lvwProjects_KeyUp(KeyCode As Integer, Shift As Integer)
'    lvwProjects_Click
'End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    SaveData
End Sub

Public Sub PopulateValues(intProjectID As Integer)
Dim rst As ADODB.Recordset
On Error GoTo Catch
    SQL_SELECT
    SQLText "ProjectName"
    SQLText "ProjectPath"
    SQLText "ProjectFile"
    SQLText "CreatedDate"
    SQLText "CreatedAuthor", False
    'SQLText "ModifiedDate"
    'SQLText "ModifiedEditor"
    'SQLText "Active", False
    SQL_FROM "Project"
    SQL_WHERE_Integer "ID", intProjectID
    MasterOpenDB
    Set rst = MasterOpenRS(gstrSQL)
    ''''''''''''''''''ResetFields
    If Not rst.EOF Then
        ' Room Details
        'lblProjectID.Caption = intProjectID
        If rst!ProjectName <> "" Then
            txtProjectName.Text = rst!ProjectName
        Else
            txtProjectName.Text = ""
        End If
        If rst!Projectpath <> "" Then
            txtProjectPath.Text = rst!Projectpath
        Else
            txtProjectPath.Text = ""
        End If
        If rst!ProjectFile <> "" Then
            txtProjectFile.Text = rst!ProjectFile
        Else
            txtProjectFile.Text = ""
        End If
        'If rst("Active").Value = True Then
        '    chkActive.Value = vbChecked
        'Else
        '    chkActive.Value = vbUnchecked
        'End If
        ' Record Details
        'If rst("CreatedDate").Value <> "" Then
        '    lblCreatedDate.Caption = FormatDateAndTime(rst("CreatedDate").Value)
        'Else
        '    lblCreatedDate.Caption = ""
        'End If
        'If rst("CreatedAuthor").Value <> "" Then
        '    lblCreatedBy.Caption = rst("CreatedAuthor").Value
        'Else
        '    lblCreatedBy.Caption = ""
        'End If
        'If rst("ModifiedDate").Value <> "" Then
        '    lblLastModifiedDate.Caption = FormatDateAndTime(rst("ModifiedDate").Value)
        'Else
        '    lblLastModifiedDate.Caption = "<Never>"
        'End If
        'If rst("ModifiedEditor").Value <> "" Then
        '    lblLastModifiedBy.Caption = rst("ModifiedEditor").Value
        'Else
        '    lblLastModifiedBy.Caption = "<Never>"
        'End If
    Else
        ' Project Details
        'lblProjectID.Caption = intProjectID
        txtProjectName.Text = ""
        txtProjectPath.Text = ""
        txtProjectFile.Text = ""
        'chkActive.Value = vbChecked
        ' Record Details
        'lblCreatedDate.Caption = "<New>"
        'lblCreatedBy.Caption = "<New>"
        'lblLastModifiedDate.Caption = "<New>"
        'lblLastModifiedBy.Caption = "<New>"
    End If
    CloseRS rst
    MasterCloseDB
    Exit Sub
Catch:
    CloseRS rst
    MasterCloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "PopulateValues"
    LogError "Error", "PopulateValues", Err.Description
End Sub

'Private Sub ListProjects()
'Dim rst As ADODB.Recordset
'Dim List As ListItem
'Dim LSI As ListSubItem
'Dim i As Integer
'On Error GoTo Catch
'    lvwProjects.ListItems.Clear
'    SQL_SELECT
'    SQLText "ID"
'    SQLText "ProjectName"
'    SQLText "ProjectPath"
'    SQLText "ProjectFile", False
'    'SQLText "Active", False
'    SQL_FROM "Project"
'    MasterOpenDB
'    Set rst = MasterOpenRS(gstrSQL)
'    While Not rst.EOF
'        Set List = lvwProjects.ListItems.Add(, "i" & rst!ID, rst!ID, 0, 0)
'        If rst!ProjectName <> "" Then
'            List.SubItems(1) = rst!ProjectName
'        Else
'            List.SubItems(1) = ""
'        End If
'        'List.SubItems(2) = rst!RoomLongName
'        If rst!Projectpath <> "" Then
'            List.SubItems(2) = rst!Projectpath
'        Else
'            List.SubItems(2) = ""
'        End If
'        If rst!ProjectFile <> "" Then
'            List.SubItems(3) = rst!ProjectFile
'        Else
'            List.SubItems(3) = ""
'        End If
'        lvwProjects.Refresh
'        rst.MoveNext
'    Wend
'    CloseRS rst
'    MasterCloseDB
'    lvwProjects_Click
'    Exit Sub
'Catch:
'    CloseRS rst
'    MasterCloseDB
'    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "ListProjects"
'    LogError "Error", "ListProjects", Err.Description
'End Sub

Private Sub PopulateData()
    txtProjectName.Text = gstrProjectName
    txtProjectPath.Text = gstrProjectPath
    txtProjectFile.Text = gstrProjectFile
End Sub

Private Sub SaveData()
Dim rst As ADODB.Recordset
Dim intProjectID As Integer
Dim lngRecordsAffected As Long
On Error GoTo Catch
    If vbNo = MsgBox("Do you want to Save this Project?", vbQuestion + vbYesNo, "SaveData") Then
        Exit Sub
    End If
    'intProjectID = CInt(lblProjectID.Caption)
    'intProjectID = lvwProjects.SelectedItem.Text
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
    If Trim(txtProjectFile.Text) = "" Then
        MsgBox "Please key in Project File", vbExclamation, "SaveData"
        txtProjectFile.SetFocus
        Exit Sub
    End If
    SQL_SELECT_ALL "Project"
    'SQL_WHERE_Integer "ID", intProjectID
    SQL_WHERE_Text "ProjectName", gstrProjectName
    SQL_AND_Text "ProjectPath", gstrProjectPath
    SQL_AND_Text "ProjectFile", gstrProjectFile
    
    MasterOpenDB
    Set rst = MasterOpenRS(gstrSQL)
    If Not rst.EOF Then
        'Update table Project
        SQL_UPDATE "Project"
        SQL_SET_Text "ProjectName", Trim(txtProjectName.Text)
        SQL_SET_Text "ProjectPath", Trim(txtProjectPath.Text)
        SQL_SET_Text "ProjectFile", Trim(txtProjectFile.Text)
        ' Record Details
        SQL_SET_DateTime "ModifiedDate", Now
        SQL_SET_Text "ModifiedEditor", "Aeric", False
        'If chkActive.Value = vbChecked Then
        '    SQL_SET_Boolean "Active", True, False
        'Else
        '    SQL_SET_Boolean "Active", False, False
        'End If
        SQL_WHERE_Integer "ID", intProjectID
    Else
'        SQL_INSERT "Project"
'        'SQLText "ID"
'        SQLText "ProjectName"
'        SQLText "ProjectPath"
'        SQLText "ProjectFile"
'        SQLText "CreatedDate"
'        SQLText "CreatedAuthor", False
'        SQL_VALUES
'        SQLData_Text Trim(txtProjectName.Text)
'        SQLData_Text Trim(txtProjectPath.Text)
'        SQLData_Text Trim(txtProjectFile.Text)
'        SQLData_DateTime Now
'        SQLData_Text "Aeric", False
'        SQL_Close_Bracket
        
        MsgBox "Project not found!", vbExclamation, "SaveData"
        CloseRS rst
        MasterCloseDB
        Exit Sub
    End If
    CloseRS rst
    MasterQuerySQL gstrSQL, lngRecordsAffected
    MasterCloseDB
    MsgBox "Project is saved!", vbInformation, "SaveData"
    'ResetFields
    'ListProjects
    'SelectProject intProjectID
    'lvwProjects_Click
    gstrProjectName = Trim(txtProjectName.Text)
    gstrProjectPath = Trim(txtProjectPath.Text)
    gstrProjectFile = Trim(txtProjectFile.Text)
    Exit Sub
Catch:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "SaveData"
    LogError "Error", "SaveData", Err.Description
    CloseRS rst
    MasterCloseDB
End Sub

'Public Sub SelectProject(pintProjectID As Integer)
'Dim i As Integer
'Dim blnExist As Boolean
'On Error GoTo Catch
'    With lvwProjects
'        For i = 1 To .ListItems.Count
'            If .ListItems.Item(i).Text = pintProjectID Then
'                '.ListIndex = i
'                .ListItems(i).Selected = True
'                .SelectedItem.EnsureVisible
'                blnExist = True
'                lvwProjects_Click
'                Exit For
'            End If
'        Next
'        If blnExist = False Then
'            'lblRoomID.Caption = pintProjectID ' ""
'            ResetFields
'        End If
'    End With
'    Exit Sub
'Catch:
'    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "SelectProject"
'    LogError "Error", "SelectProject", Err.Description
'End Sub

'Private Sub ResetFields()
'    txtProjectName.Text = ""
'    txtProjectPath.Text = ""
'    txtProjectFile.Text = ""
'End Sub
'
'Private Sub mnuFileNew_Click()
'    ResetFields
'End Sub
