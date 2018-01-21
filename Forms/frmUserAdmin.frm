VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserAdmin 
   Caption         =   "User Admin"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   11325
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList imageList1 
      Left            =   10440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserAdmin.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrTop 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "imageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "EXIT"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwUsers 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Column1"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Column2"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Column3"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Column4"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Admin"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5160
      TabIndex        =   1
      Top             =   840
      Width           =   2520
   End
   Begin VB.Shape shpTop 
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   600
      Width           =   11295
   End
End
Attribute VB_Name = "frmUserAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
'Dim pstrProjectName As String
'Dim pstrProjectFolder As String
'Dim pstrProjectPath As String
'Dim pstrProjectData As String
'Dim pstrProjectDataPath As String
'Dim pstrProjectDataFile As String
'Dim pstrProjectDataPassword As String
'Dim pstrProjectItemsFile As String
'Dim pstrProjectItemsPassword As String

Private Sub Form_Load()
    InitGlobalVariables
    ListUser
End Sub

Private Sub tbrTop_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "EXIT"
            Unload Me
    End Select
End Sub

' Initialize project level variables
'Private Sub InitProjectVariables()
'    pstrProjectName = ""
'
'    pstrProjectFolder = "\Projects\" ' default
'    pstrProjectPath = App.Path & gstrProjectFolder
'    pstrProjectData = "\Data\"
'    pstrProjectDataPath = gstrProjectPath & gstrProjectData
'    pstrProjectDataFile = "Data.mdb"
'    pstrProjectDataPassword = ""
'    pstrProjectItemsFile = "Items.mdb"
'    pstrProjectItemsPassword = ""
'
''    gstrMasterFolder = "\"
''    gstrMasterPath = App.Path & gstrMasterFolder
''    gstrMasterData = "Databases\"
''    gstrMasterDataPath = gstrMasterPath & gstrMasterData
''    gstrMasterDataFile = "Projects.mdb"
''    gstrMasterDataPassword = ""
'End Sub

' Initialize global variables
Private Sub InitGlobalVariables()
    'gstrProjectName = ""
    'gstrProjectFile = ""
    
    'gstrProjectFolder = "\Projects\" ' default
    gstrProjectPath = App.Path '& gstrProjectFolder
    gstrProjectData = "\Data\"
    gstrProjectDataPath = gstrProjectPath & gstrProjectData
    gstrProjectDataFile = "Data.mdb"
    gstrProjectDataPassword = ""
    gstrProjectItemsFile = "Items.mdb"
    gstrProjectItemsPassword = ""
    
'    gstrMasterFolder = "\"
'    gstrMasterPath = App.Path & gstrMasterFolder
'    gstrMasterData = "Databases\"
'    gstrMasterDataPath = gstrMasterPath & gstrMasterData
'    gstrMasterDataFile = "Projects.mdb"
'    gstrMasterDataPassword = ""
End Sub

Private Sub ListUser()
Dim rst As ADODB.Recordset
Dim List As ListItem
Dim LSI As ListSubItem
Dim i As Integer
On Error GoTo Catch
    lvwUsers.ListItems.Clear
    SQL_SELECT
    SQLText "ID"
    SQLText "UserName"
    SQLText "FirstName"
    SQLText "LastName"
    SQLText "Active", False
    SQL_FROM "[User]"
    ProjectOpenDB
    Set rst = ProjectOpenRS(gstrSQL)
    While Not rst.EOF
        Set List = lvwUsers.ListItems.Add(, "i" & rst!ID, rst!ID, 0, 0)
        If rst!UserName <> "" Then
            List.SubItems(1) = rst!UserName
        Else
            List.SubItems(1) = ""
        End If
        If rst!FirstName <> "" Then
            List.SubItems(2) = rst!FirstName
        Else
            List.SubItems(2) = ""
        End If
        If rst!LastName <> "" Then
            List.SubItems(3) = rst!LastName
        Else
            List.SubItems(3) = ""
        End If
        If rst!Active = True Then
            List.SubItems(4) = "Yes"
        Else
            List.SubItems(4) = "No"
        End If
        If rst!Active = False Then
            For i = 1 To 4
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = vbBlue
            Next
        Else
            For i = 1 To 4
                Set LSI = List.ListSubItems(i)
                LSI.ForeColor = vbBlack
            Next
        End If
        'With lvwUsers
            '.ColumnHeaders(1).Text = "Username"
            '.ColumnHeaders(2).Text = "Firstname"
            '.ColumnHeaders(3).Text = "Lastname"
            '.ColumnHeaders(4).Text = "Active"
            '.Refresh
        'End With
        rst.MoveNext
    Wend
    CloseRS rst
    ProjectCloseDB
    With lvwUsers
        .ColumnHeaders(2).Text = "Username"
        .ColumnHeaders(3).Text = "Firstname"
        .ColumnHeaders(4).Text = "Lastname"
        .ColumnHeaders(5).Text = "Active"
        .Refresh
    End With
    'lvwUsers_Click
    Exit Sub
Catch:
    CloseRS rst
    ProjectCloseDB
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "ListUsers"
    LogError "Error", "ListUsers", Err.Description
End Sub
