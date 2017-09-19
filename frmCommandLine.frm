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
Dim strViewText As String
Dim strInput As String
Private strProjectType As String
Private strProjectFolder As String
Private strProjectName As String

Private Sub Form_Load()
    txtView.Text = ""
    txtEdit.Text = ""
    'gstrDatabasePath = App.Path & "\Databases\Access"
    'gstrDatabaseFile = "project.mdb"
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        GoTo ExecuteCommand
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
                                            Else
                                                If strProjectName = "" Then
                                                    SetViewText "No active project!"
                                                    ShowCommand strViewText
                                                    CheckCommand = True
                                                    Exit Function
                                                End If
                                                If strProjectFolder = "" Then strProjectFolder = App.Path & "\Projects\"
                                                If Dir(strProjectFolder & strProjectName & "\Data", vbDirectory) <> "" Then
                                                
                                                Else
                                                    MkDir strProjectFolder & strProjectName & "\Data"
                                                End If
                                                gstrDatabasePath = strProjectFolder & strProjectName & "\Data"
                                                gstrDatabaseFile = "Data.mdb"
                                                If Not FileExists(gstrDatabasePath & "\" & gstrDatabaseFile) Then
                                                    SetViewText "Scaffolding..."
                                                    ShowCommand strViewText
                                                    If Not CreateDatabase Then
                                                        SetViewText "Database create failed!"
                                                        ShowCommand strViewText
                                                        CheckCommand = True
                                                    End If
                                                End If
                                                If CreateModel(Format(strCommand(3), vbProperCase)) Then ' [User]
                                                    SetViewText "Model created successful!"
                                                    ShowCommand strViewText
                                                    CheckCommand = True
                                                Else
                                                    SetViewText "Wrong command arguments! Do not use space for Model."
                                                    ShowCommand strViewText
                                                    CheckCommand = True
                                                End If
                                            End If
                                        Else
                                            CheckCommand = False
                                            Exit Function
                                        End If
                                    Case Else
                                        CheckCommand = False
                                        Exit Function
                                End Select
                            End If
                        Case "create"
                            If UBound(strCommand) > 1 Then
                                'If UBound(strCommand) > 2 Then
                                '    SetViewText "Wrong command arguments! Do not use space for Project_Name."
                                '    ShowCommand strViewText
                                '    CheckCommand = True
                                'Else
                                    strProjectFolder = App.Path & "\Projects\"
                                    'strProjectName = Format(strCommand(2), vbProperCase)
                                    strProjectName = Right$(strValue, Len(strValue) - Len("vb create ")) ' "vb  create" 2 spaces or more in between is invalid
                                    strProjectName = Trim(Format$(strProjectName, vbProperCase))
                                    If strProjectName = "" Then
                                        SetViewText "Invalid Project Name!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                    If isProjectExist(strProjectName, "STANDARD", strProjectFolder) Then
                                        SetViewText "Project already exist!"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    Else
                                        SetViewText "Project created successfully"
                                        ShowCommand strViewText
                                        CheckCommand = True
                                        Exit Function
                                    End If
                                'End If
                            End If
                        Case "open"
                            If UBound(strCommand) > 1 Then
                                strProjectFolder = App.Path & "\Projects\"
                                strProjectName = Right$(strValue, Len(strValue) - Len("vb open ")) ' "vb  open" 2 spaces or more in between is invalid
                                strProjectName = Trim(Format$(strProjectName, vbProperCase))
                                If strProjectName = "" Then
                                    SetViewText "Invalid Project Name!"
                                    ShowCommand strViewText
                                    CheckCommand = True
                                    Exit Function
                                End If
                                If Not FileExists(strProjectFolder & strProjectName & "\" & strProjectName & ".vbp") Then
                                    SetViewText "Project not found!"
                                    ShowCommand strViewText
                                    strProjectName = ""
                                    CheckCommand = True
                                    Exit Function
                                Else
                                    SetViewText "Project opened successfully!"
                                    ShowCommand strViewText
                                    CheckCommand = True
                                    Exit Function
                                End If
                            End If
                        Case Else
                            CheckCommand = False
                            Exit Function
                    End Select
                End If
            Case "help"
                strViewText = "Type 'vb create project_name' to create a new project"
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
End Sub

Private Sub SetViewText(strText As String)
    strViewText = strText
End Sub
