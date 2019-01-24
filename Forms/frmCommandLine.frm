VERSION 5.00
Begin VB.Form frmCommandLine 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Command Line"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14220
   Icon            =   "frmCommandLine.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8040
      Width           =   14175
   End
   Begin VB.TextBox txtView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   8055
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmCommandLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strInput As String
Private strViewText As String
Private strProjectType As String
Private strProjectName As String
Private strProjectPath As String
Private strFormName As String
Private strFormFile As String

Private Sub Form_Load()
    Dim strMessage As String
    strMessage = "====================="
    strMessage = strMessage & vbCrLf & "Omelette Command Line"
    strMessage = strMessage & vbCrLf & "====================="
    strMessage = strMessage & vbCrLf & "Type help to start"
    txtView.Text = strMessage
    txtEdit.Text = ""
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GoTo ExecuteCommand
    End If
    If KeyCode = vbKeyUp Then
        txtEdit.Text = strInput & " "
        txtEdit.SelStart = Len(txtEdit.Text)   '+ 1
        txtEdit.SelLength = 0
    End If
    Exit Sub
ExecuteCommand:
    strInput = txtEdit.Text
    txtEdit.Text = ""
    txtEdit.SelStart = 0
    txtEdit.SelLength = 0
    SetViewText strInput
    ShowCommand "CMD> " & strViewText
    If Not CheckCommand(strInput) Then
        SetViewText "Unknown command! Please type HELP for tips"
        ShowCommand strViewText
    End If
End Sub

Private Function CheckCommand(strValue As String) As Boolean
    Dim strCommand() As String
    Dim strHelp As String
    strValue = Trim(strValue)
    strCommand = Split(strValue, " ")
    If UBound(strCommand) >= 0 Then
        Select Case LCase(strCommand(0))
            Case "vb"
                If UBound(strCommand) > 0 Then
                    Select Case LCase(strCommand(1))
                        Case "generate", "gen", "/g", "-g"
                            If UBound(strCommand) > 1 Then
                                Select Case LCase(strCommand(2))
                                    Case "model", "/m", "-m"
                                        If UBound(strCommand) > 2 Then
                                            If UBound(strCommand) > 3 Then
                                                SetViewText "Wrong command arguments! Do not use space for Model."
                                                ShowCommand strViewText
                                                CheckCommand = True
                                                Exit Function
                                            Else
                                                ' Modified on: 04 Jun 2018
                                                ' Aeric: Use gstrProjectName instead of strProjectName
                                                If gstrProjectName = "" Then
                                                    SetViewText "No active project!"
                                                    ShowCommand strViewText
                                                    CheckCommand = True
                                                    Exit Function
                                                End If
                                                SetDefaultValue gstrProjectDataPath, gstrMasterPath & "\" & gstrProjectFolder & "\" & gstrProjectName & "\" & gstrProjectData
                                                CreateDirectory gstrProjectDataPath
                                                If Not FileExists(gstrProjectDataPath & "\" & gstrProjectDataFile) Then
                                                    SetViewText "Scaffolding..."
                                                    ShowCommand strViewText
                                                    If Not CreateProjectData Then
                                                        SetViewText "Database creation failed!"
                                                        ShowCommand strViewText
                                                        CheckCommand = True
                                                        Exit Function
                                                    End If
                                                End If
                                                ' Disallow model = 'user'
                                                Dim strModel As String
                                                'strModel = strCommand(3)
                                                strModel = Format(strCommand(3), vbProperCase)
                                                If LCase(strModel) = "users" Or LCase(strModel) = "user" Then ' Or LCase(strCommand(3)) = "user" Then
                                                    SetViewText "Command not allowed. Please use command: vb generate users"
                                                    ShowCommand strViewText
                                                    CheckCommand = True
                                                    Exit Function
                                                End If
                                                ' todo: check is table already exist?
                                                If FindTable(strModel) Then
                                                    SetViewText "Table (" & strModel & ") already exists!"
                                                    ShowCommand strViewText
                                                    CheckCommand = True
                                                    Exit Function
                                                End If
                                                If CreateModel(strModel) Then
                                                    SetViewText "Model (" & strModel & ") created successful!"
                                                    ShowCommand strViewText
                                                    CheckCommand = True
                                                Else
                                                    SetViewText Err.Description  '"Error when creating Model."
                                                    ShowCommand strViewText
                                                    CheckCommand = True
                                                    Exit Function
                                                End If
                                            End If
                                        Else
                                            CheckCommand = False
                                            Exit Function
                                        End If
                                    Case "users"
                                        If UBound(strCommand) > 2 Then
                                            SetViewText "Wrong command arguments!"
                                            ShowCommand strViewText
                                            CheckCommand = True
                                            Exit Function
                                        End If
                                        If gstrProjectName = "" Then
                                            SetViewText "No active project!"
                                            ShowCommand strViewText
                                            CheckCommand = True
                                            Exit Function
                                        End If
                                        'If FindTable("Users") Then
                                        '    SetViewText "Table (Users) already exists!"
                                        '    ShowCommand strViewText
                                        '    CheckCommand = True
                                        '    Exit Function
                                        'End If
                                        If Not CreateUsers("Users") Then
                                            SetViewText "Users model creation failed!" & vbCrLf & "Error: " & Err.Description
                                            ShowCommand strViewText
                                            CheckCommand = True
                                            Exit Function
                                        Else
                                            SetViewText "User model created successfully!"
                                            ShowCommand strViewText
                                            CheckCommand = True
                                            Exit Function
                                        End If
                                    Case Else
                                        CheckCommand = False
                                        Exit Function
                                End Select
                            End If
                        Case "create", "/c", "-c"
                            If UBound(strCommand) > 1 Then
                                If UBound(strCommand) > 2 Then
                                    SetViewText "Wrong command arguments! Do not use space for Project_Name."
                                    ShowCommand strViewText
                                    CheckCommand = True
                                Else
                                    strProjectName = Replace(strValue, "vb " & strCommand(1) & " ", "")
                                    If strProjectName = "" Then
                                        SetViewText "Invalid Project Name!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    ' Check Valid Name (Make sure start with a to z)
                                    If Not ValidateName(strProjectName) Then
                                        SetViewText "Invalid Project Name!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    'strProjectName = Format$(Trim(strProjectName), vbProperCase)
                                    If ProjectExist(strProjectName) Then
                                        SetViewText "Project already exist!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    ' Create a "STANDARD" project
                                    If Not CreateStandardProject(strProjectName) Then
                                        SetViewText "Project creation failed!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    ' Save generated file list into a database file
                                    ' todo: Check Data folder exist and create if not
                                    ' todo: Generate File.mdb with tables (from code or copy from a template ?)
                                    ' Update Master
                                    If Not AddProject(gstrProjectName, gstrProjectPath, gstrProjectFile, gstrProjectData) Then
                                        SetViewText "Database add new project failed"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    
                                    If Not FileExists(gstrProjectPath & "\" & gstrProjectFile) Then
                                        SetViewText "Project not found!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    Else
                                        If LoadProject Then
                                            SetViewText "Project loaded successfully!"
                                        Else
                                            SetViewText "Project failed to load!"
                                        End If
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                End If
                            End If
                        Case "drop", "/d", "-d"
                            If UBound(strCommand) > 1 Then
                                If UBound(strCommand) > 2 Then
                                    SetViewText "Wrong command arguments!"
                                    ShowCommand strViewText
                                    CheckCommand = True
                                    Exit Function
                                End If
                                Dim strTable As String
                                strTable = Right$(strValue, Len(strValue) - Len("vb " & strCommand(1) & " "))
                                If Not FindTable(strTable) Then
                                    SetViewText "Table (" & strTable & ") not found!"
                                    ShowCommand strViewText
                                    CheckCommand = True
                                    Exit Function
                                End If
                                If Not DropTable(strTable) Then
                                    SetViewText "Drop Table failed!"
                                    ShowCommand strViewText
                                    CheckCommand = True
                                    Exit Function
                                End If
                                SetViewText "Table [" & strTable & "] dropped successfully"
                                ShowCommand strViewText
                                CheckCommand = True
                                Exit Function
                            Else
                                CheckCommand = False
                                Exit Function
                            End If
                        Case "open", "/o", "-o"
                            If UBound(strCommand) > 1 Then
                                strProjectName = Right$(strValue, Len(strValue) - Len("vb " & strCommand(1) & " ")) ' "vb  open" 2 spaces or more in between is invalid
                                strProjectName = Trim(Format$(strProjectName, vbProperCase))
                                strProjectPath = gstrMasterPath & "\" & gstrProjectFolder & "\" & strProjectName
                                If strProjectName = "" Then
                                    SetViewText "Invalid Project Name!"
                                    ShowCommand strViewText
                                    CheckCommand = True
                                    Exit Function
                                End If
                                If Not FileExists(strProjectPath & "\" & strProjectName & ".vbp") Then
                                    SetViewText "Project not found!"
                                    ShowCommand strViewText
                                    'strProjectName = ""
                                    CheckCommand = True
                                    Exit Function
                                Else
                                    gstrProjectName = strProjectName
                                    gstrProjectPath = strProjectPath
                                    gstrProjectDataPath = gstrProjectPath & "\" & gstrProjectData
                                    If LoadProject Then
                                        SetViewText "Project loaded successfully!"
                                    Else
                                        SetViewText "Project failed to load!"
                                    End If
                                    ShowCommand strViewText
                                    CheckCommand = True
                                    Exit Function
                                End If
                            End If
                        Case "find", "/f", "-f"
                            If UBound(strCommand) > 1 Then
                                If gstrProjectDataPath = "" Then
                                    SetViewText "Unknown project!"
                                Else
                                    If FindTable(strCommand(2)) Then
                                        SetViewText "Table found!"
                                    Else
                                        SetViewText "Table not found!"
                                    End If
                                End If
                                ShowCommand strViewText
                                CheckCommand = True
                                Exit Function
                            End If
                        Case "build", "/b", "-b" ' ignore parameters
                            'If UBound(strCommand) > 1 Then
                            '    CheckCommand = False
                            '    Exit Function
                            'End If
                            CompileExe
                            SetViewText "Build successfully!"
                            ShowCommand strViewText
                            CheckCommand = True
                        Case "run", "/r", "-r" ' ignore parameters
                            RunExe
                            SetViewText "Running Application..."
                            ShowCommand strViewText
                            CheckCommand = True
                        Case "start", "/s", "-s" ' build and run
                            Dim strProjectExe As String
                            strProjectExe = gstrMasterPath & "\" & gstrProjectFolder & "\" & gstrProjectName & "\" & gstrProjectName & ".exe"
                            If Not FileExists(strProjectExe) Then
                                CompileExe
                            End If
                            RunExe
                            SetViewText "Application started!"
                            ShowCommand strViewText
                            CheckCommand = True
                        Case Else
                            CheckCommand = False
                            'mdiMain.sbStatus.Panels(1).Text = ""
                            Exit Function
                    End Select
                End If
            Case "pwd"
                strViewText = "Current project: " & gstrProjectName
                ShowCommand strViewText
                CheckCommand = True
                Exit Function
            Case "help"
                strHelp = vbCrLf & "[ Sample Commands ]"
                strHelp = strHelp & vbCrLf & "vb create project_name" & vbTab & vbTab & "to create a new project"
                strHelp = strHelp & vbCrLf & "vb open project_name" & vbTab & vbTab & "to open existing project (inside Projects folder)"
                strHelp = strHelp & vbCrLf & "vb build" & vbTab & vbTab & vbTab & vbTab & "to compile opened project into exe"
                strHelp = strHelp & vbCrLf & "vb run" & vbTab & vbTab & vbTab & vbTab & "to run compiled exe"
                strHelp = strHelp & vbCrLf & "vb generate users" & vbTab & vbTab & vbTab & "to create a default Users table"
                strHelp = strHelp & vbCrLf & "vb generate model model_name" & vbTab & "to create a database with table name model_name"
                'strHelp = strHelp & vbCrLf & "vb g m model_name" & vbTab & vbTab & vbTab & "same as 'vb generate model model_name'"
                'strHelp = strHelp & vbCrLf & "vb list model" & vbTab & vbTab & vbTab & "list all models inside the database"
                strHelp = strHelp & vbCrLf & "help" & vbTab & vbTab & vbTab & vbTab & "to display this message"
                strHelp = strHelp & vbCrLf & "exit / quit" & vbTab & vbTab & vbTab & vbTab & "close this window"
                strViewText = strHelp
                ShowCommand strViewText
                CheckCommand = True
                Exit Function
            Case "exit", "quit", "/q"
                strViewText = "Closing..."
                ShowCommand strViewText
                CheckCommand = True
                Unload Me
                Exit Function
            Case Else
                CheckCommand = False
                Exit Function
        End Select
    End If
End Function

Private Sub ShowCommand(strText As String)
    txtView.Text = txtView.Text & vbCrLf & strText
    txtView.SelStart = Len(txtView.Text)
    txtView.SelLength = 0
End Sub

Private Sub SetViewText(strText As String)
    strViewText = strText
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case vbKeyUp
'            txtEdit.Text = strInput
'        Case Else
'
'    End Select
End Sub

Private Function CreateStandardProject(ByVal strProjectName As String) As Boolean
Dim strProjectDataPath As String
Dim strProjectFile As String
Try:
    On Error GoTo Catch
    SetDefaultValue strProjectName, "Project1", True
    strProjectName = Format$(strProjectName, vbProperCase)
    strProjectFile = strProjectName & ".vbp"
    strProjectPath = gstrMasterPath & "\" & gstrProjectFolder & "\" & strProjectName
    CreateDirectory strProjectPath
    gstrProjectDataPath = ""
    strProjectDataPath = strProjectPath & "\" & gstrProjectData
    CreateDirectory strProjectDataPath
    'gstrProjectItemsFile = "Items.mdb"
    If Not FileExists(strProjectDataPath & gstrProjectItemsFile) Then
        SetViewText "Creating database (Items)..."
        ShowCommand strViewText
        If Not CreateProjectItems(strProjectDataPath) Then
            SetViewText "Database (Items) creation failed!"
            ShowCommand strViewText
            CreateStandardProject = False
            Exit Function
        End If
        ' [Project]
        If Not CreateItems("Project", "Project", True, strProjectDataPath) Then
            SetViewText "Table (Project) creation failed!"
            ShowCommand strViewText
            CreateStandardProject = False
            Exit Function
        End If
        SetViewText "Table (Project) created successful..."
        ShowCommand strViewText
        
        If Not WriteVbp("STANDARD", strProjectPath, strProjectFile, strProjectName) Then
            SetViewText "Write file (Project) failed!"
            ShowCommand strViewText
            CreateStandardProject = False
            Exit Function
        End If
        SetViewText "File (" & strProjectFile & ") created successful..."
        ShowCommand strViewText
        If Not AddItem("Project", strProjectName, strProjectPath, strProjectFile, "Storage", strProjectName) Then
            SetViewText "Update (Project) table failed!" & " " & Err.Description
            ShowCommand strViewText
            CreateStandardProject = False
            Exit Function
        End If
        SetViewText "Table (Project) updated successful..."
        ShowCommand strViewText
        ' [Forms]
        If Not CreateItems("Forms", "Form", True, strProjectDataPath) Then
            SetViewText "Table (Forms) creation failed!"
            ShowCommand strViewText
            CreateStandardProject = False
            Exit Function
        End If
        SetViewText "Table (Forms) created successful..."
        ShowCommand strViewText
        strFormName = "frmForm1"
        strFormFile = "frmForm1.frm"
        If Not WriteFrm("STANDARD", strProjectPath, strFormName) Then
            SetViewText "Write file (Form) failed!"
            ShowCommand strViewText
            CreateStandardProject = False
            Exit Function
        End If
        SetViewText "File (" & strFormFile & ") created successful..."
        ShowCommand strViewText
        If Not AddItem("Forms", strFormName, strProjectPath, strFormFile, "Storage", strProjectName) Then
            SetViewText "Update (Forms) table failed!"
            ShowCommand strViewText
            CreateStandardProject = False
            Exit Function
        End If
        SetViewText "Table (Forms) updated successful..."
        ShowCommand strViewText
        ' [Modules]
        If Not CreateItems("Modules", "Module", True, strProjectDataPath) Then
            SetViewText "Table (Modules) creation failed!"
            ShowCommand strViewText
            CreateStandardProject = False
            Exit Function
        End If
        SetViewText "Table (Modules) created successful..."
        ShowCommand strViewText
        ' [Class]
        If Not CreateItems("Class", "Class", True, strProjectDataPath) Then
            SetViewText "Table (Class) creation failed!"
            ShowCommand strViewText
            CreateStandardProject = False
            Exit Function
        End If
        SetViewText "Table (Class) created successful..."
        ShowCommand strViewText
    End If
    gstrProjectName = strProjectName
    gstrProjectFile = gstrProjectName & ".vbp"
    gstrProjectPath = gstrMasterPath & "\" & gstrProjectFolder & "\" & gstrProjectName
    'gstrProjectDataFile = "Data.mdb"
    'gstrProjectItemsFile = "Items.mdb"
    gstrProjectDataPath = gstrProjectPath & "\" & gstrProjectData
    CreateStandardProject = True
    Exit Function
Catch:
    SetViewText "Unhandled error in CreateStandardProject!"
    ShowCommand strViewText
    CreateStandardProject = False
End Function

Private Function LoadProject() As Boolean
Try:
    On Error GoTo Catch
    With mdiMain
        .Caption = App.Title & " - [" & gstrProjectName & "]"
        .sbStatus.Panels(1).Text = gstrProjectPath
        .mnuFileManageProject.Enabled = True
        .mnuFileMakeExe.Caption = "&Compile " & gstrProjectName & ".exe..."
        .mnuFileMakeExeAndRun.Caption = "Compile and &Run " & gstrProjectName & ".exe..."
        .mnuFileMakeExe.Enabled = True
        .mnuFileMakeExeAndRun.Enabled = True
        With .tbrMain
            .Buttons(3).Enabled = True
            .Buttons(4).Enabled = True
            .Buttons(5).Enabled = True
            .Buttons(6).Enabled = True
        End With
    End With
    Unload frmWindowProject
    With frmWindowProject
        .Width = 4000
        .Height = 4000
        .Left = mdiMain.Width - .Width - 300
        .Top = 0
        .Show
    End With
    LoadProject = True
    Exit Function
Catch:
    LoadProject = False
End Function

Private Function ValidateName(ByVal strName As String) As Boolean
    Dim i As Integer
    Dim c As String
    strName = LCase(strName)
    c = Left$(strName, 1)
    If Asc(c) < 97 Or Asc(c) > 122 Then ' Not starts with a or z
        ValidateName = False
        Exit Function
    End If
    For i = 1 To Len(strName) - 1 ' 2nd chars onwards
        c = Mid(strName, i, 1)
        If Asc(c) < 97 Or Asc(c) > 122 Then ' Not starts with a or z
            If Not (Asc(c) > 47 And Asc(c) < 58) Then ' Number 0 to 9
                If Asc(c) <> 95 Then ' underscore(95)  // space(32)
                    ValidateName = False
                    Exit Function
                End If
            End If
        End If
    Next
    ValidateName = True
End Function
