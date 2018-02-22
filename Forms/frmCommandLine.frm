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
'Private strProjectFolder As String
Private strProjectPath As String
Private strFormName As String
Private strFormFile As String

Private Sub Form_Load()
    Dim strMessage As String
    strMessage = "Omelette Command Line"
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
    ShowCommand strViewText
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
                        Case "generate"
                            If UBound(strCommand) > 1 Then
                                Select Case LCase(strCommand(2))
                                    Case "model"
                                        If UBound(strCommand) > 2 Then
                                            If UBound(strCommand) > 3 Then
                                                SetViewText "Wrong command arguments! Do not use space for Model."
                                                ShowCommand strViewText
                                                CheckCommand = True
                                                Exit Function
                                            Else
                                                If strProjectName = "" Then
                                                    SetViewText "No active project!"
                                                    ShowCommand strViewText
                                                    CheckCommand = True
                                                    Exit Function
                                                End If
                                                Debug.Print strProjectName
                                                'Debug.Print strProjectFolder
                                                gstrProjectDataPath = gstrProjectPath & "\" & strProjectName & "\" & gstrProjectData
                                                If Dir(gstrProjectDataPath, vbDirectory) <> "" Then
                                                
                                                Else
                                                    MkDir gstrProjectDataPath
                                                End If
                                                gstrProjectName = strProjectName
                                                If Not FileExists(gstrProjectDataPath & "\" & gstrProjectDataFile) Then
                                                    SetViewText "Scaffolding..."
                                                    ShowCommand strViewText
                                                    If Not ProjectCreateDatabase Then
                                                        SetViewText "Database create failed!"
                                                        ShowCommand strViewText
                                                        CheckCommand = True
                                                        Exit Function
                                                    End If
                                                End If
                                                ' Disallow model = 'user'
                                                If LCase(strCommand(3)) = "user" Then
                                                    SetViewText "Command not allowed. Please use command: vb generate user"
                                                    ShowCommand strViewText
                                                    CheckCommand = True
                                                    Exit Function
                                                End If
                                                ' todo: check is table already exist?
                                                If CreateModel(Format(strCommand(3), vbProperCase)) Then
                                                    SetViewText "Model created successful!"
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
                                    Case "user"
                                        If UBound(strCommand) > 2 Then
                                            SetViewText "Wrong command arguments!"
                                            ShowCommand strViewText
                                            CheckCommand = True
                                            Exit Function
                                        End If
                                        If strProjectName = "" Then
                                            SetViewText "No active project!"
                                            ShowCommand strViewText
                                            CheckCommand = True
                                            Exit Function
                                        End If
                                        If Not CreateUserModel Then
                                            SetViewText "User model create failed!" & vbCrLf & "Error: " & Err.Description
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
                        Case "create"
                            If UBound(strCommand) > 1 Then
                                If UBound(strCommand) > 2 Then
                                    SetViewText "Wrong command arguments! Do not use space for Project_Name."
                                    ShowCommand strViewText
                                    CheckCommand = True
                                Else
                                    strProjectName = Right$(strValue, Len(strValue) - Len("vb create ")) ' "vb  create" 2 spaces or more in between is invalid
                                    strProjectName = Format$(Trim(strProjectName), vbProperCase)
                                    If strProjectName = "" Then
                                        SetViewText "Invalid Project Name!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    If ProjectExist(strProjectName) Then
                                        SetViewText "Project already exist!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    ' Create project
                                    If Not CreateProject(strProjectName, "STANDARD") Then
                                        SetViewText "Project create failed!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    gstrProjectName = strProjectName
                                    gstrProjectFile = strProjectName & ".vbp"
                                    gstrProjectDataPath = gstrProjectPath & "\" & gstrProjectName & "\" & gstrProjectData
                                    If Dir(gstrProjectDataPath, vbDirectory) <> "" Then
                                                
                                    Else
                                        MkDir gstrProjectDataPath
                                    End If
                                    If Not FileExists(gstrProjectDataPath & "\" & gstrProjectItemsFile) Then
                                        SetViewText "Creating project info..."
                                        ShowCommand strViewText
                                        If Not ItemsCreateDatabase Then
                                            SetViewText "Items file create failed!"
                                            ShowCommand strViewText
                                            CheckCommand = True
                                            Exit Function
                                        End If
                                    End If
                                    ' todo: check is table already exist?
                                    If CreateItemsModel("Project") Then  ' [Project]
                                        SetViewText "Project Item created successful!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                    Else
                                        SetViewText Err.Description  '"Error when creating Model."
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    If CreateItemsModel("Forms", "Form") Then ' [Forms]
                                        SetViewText "Forms Item created successful!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                    Else
                                        SetViewText Err.Description  '"Error when creating Model."
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    If CreateItemsModel("Modules", "Module") Then  ' [Modules]
                                        SetViewText "Modules Item created successful!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                    Else
                                        SetViewText Err.Description  '"Error when creating Model."
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    ' Save generated file list into a database file
                                    ' todo: Check Data folder exist and create if not
                                    ' todo: Generate File.mdb with tables (from code or copy from a template ?)
                                    If Not ItemAddNew(gstrProjectName, gstrProjectPath & "\" & gstrProjectName, gstrProjectFile, "Project") Then
                                        SetViewText "Database add new item (project) failed"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    strFormName = "frmForm1"
                                    strFormFile = "frmForm1"
                                    If Not ItemAddNew(strFormName, gstrProjectPath & "\" & gstrProjectName, strFormFile & ".frm", "Form") Then
                                        SetViewText "Database add new item (form) failed"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    ' Update Master
                                    If Not ProjectAddNew(gstrProjectName, gstrProjectPath & "\" & gstrProjectName, gstrProjectData, gstrProjectFile) Then
                                        SetViewText "Database add new project failed"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    With mdiMain
                                        .Caption = App.Title & " - [" & gstrProjectName & "]"
                                        .sbStatus.Panels(1).Text = gstrProjectName
                                        .mnuFileManageProject.Enabled = True
                                        .mnuFileMakeExe.Caption = "&Compile " & gstrProjectName & ".exe..."
                                        .mnuFileMakeExeAndRun.Caption = "Compile and &Run " & gstrProjectName & ".exe..."
                                        .mnuFileMakeExe.Enabled = True
                                        .mnuFileMakeExeAndRun.Enabled = True
                                    End With
                                    Unload frmWindowProject
                                    With frmWindowProject
                                        .Width = 4000
                                        .Height = 4000
                                        .Left = mdiMain.Width - .Width - 300
                                        .Top = 0
                                        .Show
                                        '.RefreshTreeview
                                    End With
                                    SetViewText "Project created successfully"
                                    ShowCommand strViewText
                                    CheckCommand = True
                                    Exit Function
                                End If
                            End If
                        Case "open"
                            If UBound(strCommand) > 1 Then
                                strProjectName = Right$(strValue, Len(strValue) - Len("vb open ")) ' "vb  open" 2 spaces or more in between is invalid
                                strProjectName = Trim(Format$(strProjectName, vbProperCase))
                                strProjectPath = App.Path & "\" & gstrProjectFolder
                                If strProjectName = "" Then
                                    SetViewText "Invalid Project Name!"
                                    ShowCommand strViewText
                                    CheckCommand = True
                                    'mdiMain.sbStatus.Panels(1).Text = ""
                                    Exit Function
                                End If
                                If Not FileExists(strProjectPath & "\" & strProjectName & "\" & strProjectName & ".vbp") Then
                                    SetViewText "Project not found!"
                                    ShowCommand strViewText
                                    strProjectName = ""
                                    CheckCommand = True
                                    'mdiMain.sbStatus.Panels(1).Text = ""
                                    Exit Function
                                Else
                                    'gstrProjectFolder = strProjectFolder
                                    gstrProjectName = strProjectName
                                    gstrProjectDataPath = strProjectPath & "\" & gstrProjectName & "\" & gstrProjectData
                                    gstrProjectFile = strProjectName & ".vbp"
                                    With mdiMain
                                        .Caption = App.Title & " - [" & gstrProjectName & "]"
                                        .sbStatus.Panels(1).Text = strProjectPath & "\" & gstrProjectName ' gstrProjectName
                                        .mnuFileManageProject.Enabled = True
                                        .mnuFileMakeExe.Caption = "&Compile " & gstrProjectName & ".exe..."
                                        .mnuFileMakeExeAndRun.Caption = "Compile and &Run " & gstrProjectName & ".exe..."
                                        .mnuFileMakeExe.Enabled = True
                                        .mnuFileMakeExeAndRun.Enabled = True
                                    End With
                                    Unload frmWindowProject
                                    With frmWindowProject
                                        .Width = 4000
                                        .Height = 4000
                                        .Left = mdiMain.Width - .Width - 300
                                        .Top = 0
                                        .Show
                                        '.RefreshTreeview
                                    End With
                                    SetViewText "Project opened successfully!"
                                    ShowCommand strViewText
                                    CheckCommand = True
                                    Exit Function
                                End If
                            End If
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
                strHelp = strHelp & vbCrLf & "vb generate user" & vbTab & vbTab & vbTab & "to create a default Users table"
                strHelp = strHelp & vbCrLf & "vb generate model model_name" & vbTab & "to create a database with table name model_name"
                'strHelp = strHelp & vbCrLf & "vb g m model_name" & vbTab & vbTab & vbTab & "same as 'vb generate model model_name'"
                'strHelp = strHelp & vbCrLf & "vb list model" & vbTab & vbTab & vbTab & "list all models inside the database"
                strHelp = strHelp & vbCrLf & "help" & vbTab & vbTab & vbTab & vbTab & "to display this message"
                strHelp = strHelp & vbCrLf & "exit / quit" & vbTab & vbTab & vbTab & vbTab & "close this window"
                strViewText = strHelp
                ShowCommand strViewText
                CheckCommand = True
                Exit Function
            Case "exit", "quit"
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
