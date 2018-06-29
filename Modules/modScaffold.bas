Attribute VB_Name = "modScaffold"
' Version : 0.4
' Author: Poon Yip Hoon
' Created On   : 18 Sep 2017
' Modified On  : 29 Jun 2018
' Descriptions : Scaffolding
' Dependencies:
' modProject
' modVariable
' OmlDatabase

Option Explicit

' CLASS / MODEL
' Command: vb generate model Model_Name
'Public gstrModel As String

' CREATE - READ - UPDATE - DELETE
' Reference: https://msdn.microsoft.com/en-us/vba/access-vba/articles/create-and-delete-tables-and-indexes-using-access-sql

' ID Int, Name Text, [OPTIONAL] Type Text, Active Bit, CreatedDate SmallDateTime, CreatedBy Text, ModifiedDate SmallDateTime, ModifiedBy Text

'Private strProjectType As String
'Private strProjectFolder As String
'Private strProjectName As String
'Private strProjectFormFile() As String
Private strProjectFormName() As String
Private strProjectModuleName() As String
Private strProjectClassName() As String
Private Type ProjectItem
    Name As String
    File As String
End Type

Public Function CreateModel(ByVal strModel As String, _
                            Optional ByVal strPrefix As String = "", _
                            Optional ByVal blnAppendModel As Boolean = False, _
                            Optional ByVal blnActiveDefault As Boolean = True, _
                            Optional ByVal DataPath As String = "", _
                            Optional ByVal DataFile As String = "") As Boolean
    On Error GoTo Catch
    Dim DB As New OmlDatabase
    With DB
        SetDefaultValue DataPath, gstrProjectDataPath
        SetDefaultValue DataFile, gstrProjectDataFile
        .DataPath = DataPath & "\"
        .DataFile = DataFile
        .OpenMdb
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            CreateModel = False
            Exit Function
        End If
        SQL_CREATE strPrefix & strModel
        If blnAppendModel Then
            SQL_COLUMN_ID strModel & "ID"
            SQL_COLUMN_TEXT "[" & strModel & "Name]" ' First Column
            'SQL_COLUMN_TEXT "[" & strModel & "Type]" ' Second Column
        Else
            SQL_COLUMN_ID
            SQL_COLUMN_TEXT "[Name]"
            'SQL_COLUMN_TEXT "[Type]"
        End If
        SQL_COLUMN_YESNO "[Active]"
        SQL_COLUMN_DATETIME "[CreatedDate]", "NOW()"
        SQL_COLUMN_TEXT "[CreatedBy]", 100, "Omelette"  ' Change to your name
        SQL_COLUMN_DATETIME "[ModifiedDate]"
        SQL_COLUMN_TEXT "[ModifiedBy]", 100, "", True, True, False
        SQL_Close_Bracket
        'Debug.Print gstrSQL
        .Execute gstrSQL
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            CreateModel = False
            Exit Function
        End If
        .CloseMdb
    End With
    
    'gstrSQL = "CREATE INDEX idx" & strModel & "ID"
    'gstrSQL = gstrSQL & " ON " & strPrefix & strModel & " ("
    'If blnAppendModel Then
    '    gstrSQL = gstrSQL & strModel
    'End If
    'gstrSQL = gstrSQL & "ID)"
    'gstrSQL = gstrSQL & " WITH PRIMARY"
    'OpenDB
    'QuerySQL gstrSQL
    'CloseDB
    
    'OpenDB
    'gstrSQL = "ALTER TABLE " & strPrefix & strModel
    'gstrSQL = gstrSQL & " ALTER COLUMN"
    'gstrSQL = gstrSQL & " [CreatedDate] DATETIME DEFAULT NOW() NOT NULL"
    'QuerySQL gstrSQL
    'CloseDB
    
    CreateModel = True
    Exit Function
Catch:
    Err.Description = Err.Description & " " & DB.ErrorDesc
    CreateModel = False
End Function

Public Function CreateItems(ByVal strItemModel As String, _
                            Optional ByVal strItem As String = "", _
                            Optional ByVal blnActiveDefault As Boolean = True, _
                            Optional ByVal DataPath As String = "", _
                            Optional ByVal DataFile As String = "") As Boolean
    On Error GoTo Catch
    If strItem = "" Then strItem = strItemModel
    Dim DB As New OmlDatabase
    With DB
        SetDefaultValue DataPath, gstrProjectDataPath
        SetDefaultValue DataFile, gstrProjectItemsFile
        .DataPath = DataPath & "\"
        .DataFile = DataFile
        .OpenMdb
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            CreateItems = False
            Exit Function
        End If
        SQL_CREATE strItemModel
        SQL_COLUMN_ID
        SQL_COLUMN_TEXT "[" & strItem & "Name]"
        SQL_COLUMN_TEXT "[" & strItem & "Path]"
        SQL_COLUMN_TEXT "[" & strItem & "File]"
        If strItemModel = "Project" Then
            SQL_COLUMN_TEXT "[" & strItem & "Data]" ' Default Data Folder i.e. Storage
        End If
        SQL_COLUMN_YESNO "[Active]"
        SQL_COLUMN_DATETIME "[CreatedDate]", "NOW()"
        SQL_COLUMN_TEXT "[CreatedBy]", 255, "Omelette"  ' Change to your name
        SQL_COLUMN_DATETIME "[ModifiedDate]"
        SQL_COLUMN_TEXT "[ModifiedBy]", 255, "", True, True, False
        SQL_Close_Bracket
        
        'Debug.Print gstrSQL
        .Execute gstrSQL
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            CreateItems = False
            Exit Function
        End If
        .CloseMdb
    End With
    CreateItems = True
    Exit Function
Catch:
    Err.Description = Err.Description & " " & DB.ErrorDesc
    CreateItems = False
End Function

' Created On        : 05 Dec 2017
' Last Modified On  : 29 Jun 2018
Public Function CreateUsers(Optional ByVal strModel As String = "", _
                            Optional ByVal strModelID As String = "", _
                            Optional ByVal strUserID As String = "", _
                            Optional ByVal strUserName As String = "", _
                            Optional ByVal strFirstName As String = "", _
                            Optional ByVal strLastName As String = "", _
                            Optional ByVal strUserPassword As String = "", _
                            Optional ByVal strSalt As String = "", _
                            Optional ByVal strUserRole As String = "", _
                            Optional ByVal strActiveDefault As String = "", _
                            Optional ByVal DataPath As String = "", _
                            Optional ByVal DataFile As String = "") As Boolean
    On Error GoTo Catch
    SetDefaultValue strModel, "Users"
    SetDefaultValue strModelID, "ID"
    SetDefaultValue strUserID, "UserID"
    SetDefaultValue strUserName, "UserName"
    SetDefaultValue strFirstName, "FirstName"
    SetDefaultValue strLastName, "LastName"
    SetDefaultValue strUserPassword, "UserPassword"
    SetDefaultValue strSalt, "Salt"
    SetDefaultValue strUserRole, "UserRole"
    Dim DB As New OmlDatabase
    With DB
        SetDefaultValue DataPath, gstrProjectDataPath
        SetDefaultValue DataFile, gstrProjectDataFile
        .DataPath = DataPath & "\"
        .DataFile = DataFile
        .OpenMdb
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            CreateUsers = False
            Exit Function
        End If
        
        SQL_CREATE strModel
        SQL_COLUMN_ID strModelID
        SQL_COLUMN_TEXT strUserID
        SQL_COLUMN_TEXT strUserName
        SQL_COLUMN_TEXT strFirstName, 100
        SQL_COLUMN_TEXT strLastName, 100
        SQL_COLUMN_TEXT strUserPassword
        SQL_COLUMN_TEXT strSalt
        SQL_COLUMN_TEXT strUserRole
        SQL_COLUMN_YESNO "[Active]", strActiveDefault
        SQL_COLUMN_DATETIME "[CreatedDate]", "NOW()"
        SQL_COLUMN_TEXT "[CreatedBy]", 255, "Omelette"  ' Change to your name
        SQL_COLUMN_DATETIME "[ModifiedDate]"
        SQL_COLUMN_TEXT "[ModifiedBy]", 255, "", True, True, False
        SQL_Close_Bracket
        'Debug.Print gstrSQL
        .Execute gstrSQL
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            CreateUsers = False
            Exit Function
        End If
        .CloseMdb
    End With
    CreateUsers = True
    Exit Function
Catch:
    Err.Description = Err.Description & " " & DB.ErrorDesc
    CreateUsers = False
End Function

'Public Function CreateDatabase() As Boolean
'On Error GoTo Catch
'    Dim cat As New ADOX.Catalog
'    Dim strDBCon As String
'    'strDBPath = App.Path & "\Full.lic"
'    strDBCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrDatabasePath & "\" & gstrDatabaseFile
'    strDBCon = strDBCon & ";Jet OLEDB:Database Password=" & gstrDatabasePassword
'    cat.Create strDBCon
'    CreateDatabase = True
'    Exit Function
'Catch:
'      CreateDatabase = False
'      'MsgBox Err.Description
'      Set cat = Nothing
'End Function

Public Function FindTable(ByVal strTable As String) As Boolean
Try:
    On Error GoTo Catch
    Dim DB As New OmlDatabase
    With DB
        .DataPath = gstrProjectDataPath & "\"
        .DataFile = gstrProjectDataFile
        .OpenMdb
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            FindTable = False
            Exit Function
        End If
        SQL_SELECT_ID strTable
        'Debug.Print gstrSQL
        .Execute gstrSQL
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            FindTable = False
            Exit Function
        End If
        .CloseMdb
    End With
    FindTable = True
    Exit Function
Catch:
    Err.Description = Err.Description & " " & DB.ErrorDesc
    FindTable = False
End Function

Public Function DropTable(ByVal strTable As String) As Boolean
Try:
    On Error GoTo Catch
    Dim DB As New OmlDatabase
    With DB
        .DataPath = gstrProjectDataPath & "\"
        .DataFile = gstrProjectDataFile
        .OpenMdb
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            DropTable = False
            Exit Function
        End If
        
        SQL_DROP strTable
        'Debug.Print gstrSQL
        .Execute gstrSQL
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            DropTable = False
            Exit Function
        End If
        .CloseMdb
    End With
    DropTable = True
    Exit Function
Catch:
    Err.Description = Err.Description & " " & DB.ErrorDesc
    DropTable = False
End Function

Public Function CreateProjectData(Optional DataPath As String = "", _
                                    Optional DataFile As String = "", _
                                    Optional Password As String = "") As Boolean
On Error GoTo Catch
    Dim cat As New ADOX.Catalog
    SetDefaultValue DataPath, gstrProjectDataPath
    SetDefaultValue DataFile, gstrProjectDataFile
    SetDefaultValue Password, gstrProjectDataPassword
    Dim strDBCon As String
    strDBCon = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strDBCon = strDBCon & "Data Source=" & DataPath & "\" & DataFile & ";"
    strDBCon = strDBCon & "Jet OLEDB:Database Password=" & Password
    cat.Create strDBCon
    CreateProjectData = True
    Exit Function
Catch:
    CreateProjectData = False
    Set cat = Nothing
End Function

Public Function CreateProjectItems(Optional DataPath As String = "", _
                                    Optional DataFile As String = "", _
                                    Optional Password As String = "") As Boolean
On Error GoTo Catch
    Dim cat As New ADOX.Catalog
    SetDefaultValue DataPath, gstrProjectDataPath
    SetDefaultValue DataFile, gstrProjectItemsFile
    SetDefaultValue Password, gstrProjectItemsPassword
    Dim strDBCon As String
    strDBCon = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strDBCon = strDBCon & "Data Source=" & DataPath & "\" & DataFile & ";"
    strDBCon = strDBCon & "Jet OLEDB:Database Password=" & Password
    cat.Create strDBCon
    CreateProjectItems = True
    Exit Function
Catch:
    CreateProjectItems = False
    Set cat = Nothing
End Function

Public Function AddProject(ByVal pstrProjectName As String, _
                            ByVal pstrProjectPath As String, _
                            ByVal pstrProjectFile As String, _
                            ByVal pstrProjectData As String) As Boolean
On Error GoTo Catch
    Dim DB As New OmlDatabase
    With DB
        .DataPath = gstrMasterDataPath & "\"
        .DataFile = gstrMasterDataFile
        .OpenMdb
        If .ErrorDesc <> "" Then
            AddProject = False
            Exit Function
        End If
        SQL_INSERT "Project"
        SQLText "ProjectName", True, False
        SQLText "ProjectPath"
        SQLText "ProjectFile"
        SQLText "ProjectData"
        SQLText "CreatedDate"
        SQLText "CreatedBy", False
        SQL_VALUES
        SQLData_Text pstrProjectName, True, False
        SQLData_Text pstrProjectPath
        SQLData_Text pstrProjectFile
        SQLData_Text pstrProjectData
        SQLData_DateTime Now
        SQLData_Text "Omelette", False  ' Change to your name
        SQL_Close_Bracket
        'Debug.Print gstrSQL
        .Execute gstrSQL
        If .ErrorDesc <> "" Then
            .CloseMdb
            AddProject = False
            Exit Function
        End If
        .CloseMdb
    End With
    AddProject = True
    Exit Function
Catch:
    DB.CloseMdb
    AddProject = False
End Function

Public Function AddItem(ByVal pstrItemType As String, _
                            ByVal pstrItemName As String, _
                            ByVal pstrItemPath As String, _
                            ByVal pstrItemFile As String, _
                            Optional ByVal pstrItemData As String = "Storage", _
                            Optional ByVal pstrProjectName As String = "") As Boolean
On Error GoTo Catch
    Dim DB As New OmlDatabase
    Dim strDataPath As String
    Dim strItemFile As String
    If gstrProjectDataPath = "" Then
        strDataPath = gstrMasterPath & "\" & gstrProjectFolder & "\" & pstrProjectName & "\" & gstrProjectData
    Else
        strDataPath = gstrProjectDataPath
    End If
    If gstrProjectItemsFile = "" Then
        strItemFile = "Items.mdb"
    Else
        strItemFile = gstrProjectItemsFile
    End If
    With DB
        .DataPath = strDataPath & "\"
        .DataFile = strItemFile
        .OpenMdb
        If .ErrorDesc <> "" Then
            AddItem = False
            Exit Function
        End If
        If pstrItemType = "Project" Then
            SQL_INSERT "Project"
            SQLText "ProjectName", True, False
            SQLText "ProjectPath"
            SQLText "ProjectFile"
            SQLText "ProjectData"
            SQLText "CreatedDate"
            SQLText "CreatedBy", False
            SQL_VALUES
            SQLData_Text pstrItemName, True, False
            SQLData_Text pstrItemPath
            SQLData_Text pstrItemFile
            SQLData_Text pstrItemData
            SQLData_DateTime Now
            SQLData_Text "Omelette", False  ' Change to your name
            SQL_Close_Bracket
        ElseIf pstrItemType = "Forms" Then
            SQL_INSERT "Forms"
            SQLText "FormName", True, False
            SQLText "FormPath"
            SQLText "FormFile"
            SQLText "CreatedDate"
            SQLText "CreatedBy", False
            SQL_VALUES
            SQLData_Text pstrItemName, True, False
            SQLData_Text pstrItemPath
            SQLData_Text pstrItemFile
            SQLData_DateTime Now
            SQLData_Text "Omelette", False  ' Change to your name
            SQL_Close_Bracket
        ElseIf pstrItemType = "Modules" Then
            SQL_INSERT "Modules"
            SQLText "ModuleName", True, False
            SQLText "ModulePath"
            SQLText "ModuleFile"
            SQLText "CreatedDate"
            SQLText "CreatedBy", False
            SQL_VALUES
            SQLData_Text pstrItemName, True, False
            SQLData_Text pstrItemPath
            SQLData_Text pstrItemFile
            SQLData_DateTime Now
            SQLData_Text "Omelette", False  ' Change to your name
            SQL_Close_Bracket
        ElseIf pstrItemType = "Class" Then ' Class Module
            SQL_INSERT "Class"
            SQLText "ClassName", True, False
            SQLText "ClassPath"
            SQLText "ClassFile"
            SQLText "CreatedDate"
            SQLText "CreatedBy", False
            SQL_VALUES
            SQLData_Text pstrItemName, True, False
            SQLData_Text pstrItemPath
            SQLData_Text pstrItemFile
            SQLData_DateTime Now
            SQLData_Text "Omelette", False  ' Change to your name
            SQL_Close_Bracket
        ElseIf pstrItemType = "Data" Then
            SQL_INSERT "Data"
            SQLText "DataName", True, False
            SQLText "DataPath"
            SQLText "DataFile"
            SQLText "CreatedDate"
            SQLText "CreatedBy", False
            SQL_VALUES
            SQLData_Text pstrItemName, True, False
            SQLData_Text pstrItemPath
            SQLData_Text pstrItemFile
            SQLData_DateTime Now
            SQLData_Text "Omelette", False  ' Change to your name
            SQL_Close_Bracket
        End If
        'Debug.Print gstrSQL
        .Execute gstrSQL
        If .ErrorDesc <> "" Then
            .CloseMdb
            AddItem = False
            Exit Function
        End If
        .CloseMdb
    End With
    AddItem = True
    Exit Function
Catch:
    LogError "Error", "AddItem", Err.Description
    DB.CloseMdb
    AddItem = False
End Function

' Will be replaced
'Public Function CreateProjectFiles(strProjectType As String, _
'                                strProjectPath As String, _
'                                strProjectName As String) As Boolean
'On Error GoTo Catch
'    If Trim(strProjectName) = "" Then
'        strProjectName = "Project1"
'    End If
'    If strProjectPath = "" Then strProjectPath = gstrMasterPath & "\" & gstrProjectFolder & "\" & strProjectName
'    If Dir(strProjectPath, vbDirectory) <> "" Then
'
'    Else
'        MkDir strProjectPath
'    End If
'    If FileExists(strProjectPath & "\" & strProjectName & ".vbp") Then
'
'    Else
'        If strProjectType = "DEFAULT" Then
'            WriteVbp strProjectType, strProjectPath, strProjectName & ".vbp", strProjectName
'        ElseIf strProjectType = "STANDARD" Then
'            ReDim strProjectFormName(0)
'            strProjectFormName(0) = "frmForm1.frm"
'            WriteFrm strProjectType, strProjectPath, strProjectFormName(0)
'            WriteVbp strProjectType, strProjectPath, strProjectName & ".vbp", strProjectName
'        ElseIf strProjectType = "ADMIN" Then
'            ReDim strProjectFormName(2)
'            ReDim strProjectModuleName(2)
'            ReDim strProjectClassName(0)
'            strProjectFormName(0) = "frmDashboard.frm"
'            strProjectFormName(1) = "frmLogin.frm"
'            strProjectFormName(2) = "frmStart.frm"
'            strProjectModuleName(0) = "OmeHash.bas"
'            strProjectModuleName(1) = "OmeSubMain.bas"
'            strProjectModuleName(2) = "OmeVariable.bas"
'            strProjectClassName(0) = "OmlDatabase.cls"
'            WriteFrm strProjectType, strProjectPath, strProjectFormName(0)
'            WriteFrm strProjectType, strProjectPath, strProjectFormName(1)
'            WriteFrm strProjectType, strProjectPath, strProjectFormName(2)
'            If Dir(strProjectPath & "\" & gstrProjectModules, vbDirectory) <> "" Then
'
'            Else
'                MkDir strProjectPath & "\" & gstrProjectModules
'            End If
'            WriteBas strProjectType, strProjectPath, strProjectModuleName(0)
'            WriteBas strProjectType, strProjectPath, strProjectModuleName(1)
'            WriteBas strProjectType, strProjectPath, strProjectModuleName(2)
'            If Dir(strProjectPath & "\" & gstrProjectClasses, vbDirectory) <> "" Then
'
'            Else
'                MkDir strProjectPath & "\" & gstrProjectClasses
'            End If
'            WriteCls strProjectType, strProjectPath, strProjectClassName(0)
'            WriteVbp strProjectType, strProjectPath, strProjectName & ".vbp", strProjectName
'        Else ' BLANK
'            WriteVbp "BLANK", strProjectPath, strProjectName & ".vbp", strProjectName
'        End If
'    End If
'    CreateProjectFiles = True
'    Exit Function
'Catch:
'    LogError "Error", "CreateProjectFiles", Err.Description
'    CreateProjectFiles = False
'End Function

Public Sub CreateProject(strTemplate As String, strProjectPath As String, strProjectName As String)
Try:
    On Error GoTo Catch
    Dim i As Integer
    Dim pForm() As ProjectItem
    Dim pModule() As ProjectItem
    Dim pClass() As ProjectItem
    SetDefaultValue strProjectName, "Project1", True
    strProjectName = Format$(strProjectName, vbProperCase)
    SetDefaultValue strProjectPath, gstrMasterPath & "\" & gstrProjectFolder & "\" & strProjectName
    CreateDirectory strProjectPath
    ' Create project (text files only)
'    If Not CreateProjectFiles(strTemplate, "", strProjectName) Then
'        MsgBox "Project create failed!", vbExclamation, "Error"
'        'frmProjectNew.Show
'        Exit Sub
'    End If
    gstrProjectPath = strProjectPath ' gstrMasterPath & "\" & gstrProjectFolder & "\" & strProjectName
    gstrProjectDataPath = gstrProjectPath & "\" & gstrProjectData
    CreateDirectory gstrProjectDataPath
    gstrProjectItemsFile = "Items.mdb"
    If Not FileExists(gstrProjectDataPath & gstrProjectItemsFile) Then
        ' [Project]
        If Not CreateProjectItems Then
            MsgBox "Database (Items) creation failed!", vbExclamation, "Error"
            Exit Sub
        End If
        ' todo: check is table already exist?
        If Not CreateItems("Project") Then
            MsgBox "Project Item create failed!", vbExclamation, "Error"
            Exit Sub
        End If
        ' Write out the source file
        If Not WriteVbp(strTemplate, strProjectPath, strProjectName & ".vbp", strProjectName) Then
            MsgBox "Write file (project) failed", vbExclamation, "Error"
            Exit Sub
        End If
        ' Save new item into database file
        If Not AddItem("Project", strProjectName, gstrProjectPath, strProjectName & ".vbp") Then
            MsgBox "Add new item (Project) failed", vbExclamation, "Error"
            Exit Sub
        End If
        ' [Forms]
        If Not CreateItems("Forms", "Form") Then
            MsgBox "Forms table creation failed!", vbExclamation, "Error"
            Exit Sub
        End If
        GetItems strTemplate, "Forms", pForm
        For i = 0 To UBound(pForm)
            'Debug.Print pForm(i).Name
            'Debug.Print pForm(i).File
            If Not WriteFrm(strTemplate, strProjectPath, pForm(i).Name) Then
                MsgBox "Write file (Form) failed", vbExclamation, "Error"
                Exit Sub
            End If
            If Not AddItem("Forms", pForm(i).Name, gstrProjectPath, pForm(i).File) Then
                MsgBox "Add new item (Form) failed", vbExclamation, "Error"
                Exit Sub
            End If
        Next
        ' [Modules]
        If Not CreateItems("Modules", "Module") Then
            MsgBox "Modules table creation failed!", vbExclamation, "Error"
            Exit Sub
        End If
        CreateDirectory strProjectPath & "\" & gstrProjectModules
        GetItems strTemplate, "Modules", pModule
        For i = 0 To UBound(pModule)
            If Not WriteBas(strTemplate, strProjectPath & "\" & gstrProjectModules, pModule(i).File) Then
                MsgBox "Write file (Module) failed", vbExclamation, "Error"
                Exit Sub
            End If
            If Not AddItem("Modules", pModule(i).Name, strProjectPath & "\" & gstrProjectModules, pModule(i).File) Then
                MsgBox "Add new item (Module) failed", vbExclamation, "Error"
                Exit Sub
            End If
        Next
        ' [Class]
        If Not CreateItems("Class") Then
            MsgBox "Class table creation failed!", vbExclamation, "Error"
            Exit Sub
        End If
        CreateDirectory strProjectPath & "\" & gstrProjectClasses
        GetItems strTemplate, "Class", pClass
        For i = 0 To UBound(pClass)
            If Not WriteCls(strTemplate, strProjectPath & "\" & gstrProjectClasses, pClass(i).File) Then
                MsgBox "Add new item (Class) failed", vbExclamation, "Error"
                Exit Sub
            End If
            If Not AddItem("Class", pClass(i).Name, strProjectPath & "\" & gstrProjectClasses, pClass(i).File) Then
                MsgBox "Write file (Class) failed", vbExclamation, "Error"
                Exit Sub
            End If
        Next
        'AddItemsFromTemplates strTemplate, strProjectName
    Else
        ' Items.mdb already exist
    End If
    ' [Data]
    ' Check Data file exists otherwise create one
    ' Now we can set it public/global
    SetDefaultValue gstrProjectDataPath, gstrMasterPath & "\" & gstrProjectFolder & "\" & gstrProjectName & "\" & gstrProjectData
    'If gstrProjectDataPath = "" Then gstrProjectDataPath = gstrMasterPath & "\" & gstrProjectFolder & "\" & gstrProjectName & "\" & gstrProjectData
        If Not FileExists(gstrProjectDataPath & "\" & gstrProjectDataFile) Then
            ' Create Data File
            If Not CreateProjectData Then
                MsgBox "Database (Data) creation failed!", vbExclamation, "Error"
                Exit Sub
            End If
        End If
        
        If Not CreateItems("Data") Then
            MsgBox "Data table creation failed!", vbExclamation, "Error"
            Exit Sub
        End If
        If strTemplate = "ADMIN" Then
            If Not FindTable("Users") Then
                If Not CreateUsers("Users") Then
                    MsgBox "Users model creation failed!", vbExclamation, "Error"
                    Exit Sub
                End If
            End If
            'MsgBox "User model created successfully!", vbExclamation, "Error"
        End If
    ' Update Master
    If Not AddProject(strProjectName, gstrProjectPath, strProjectName & ".vbp", gstrProjectData) Then
        MsgBox "Database add new project failed", vbExclamation, "Error"
        Exit Sub
    End If
    MsgBox "Project created successfully!", vbInformation, "Success"
        Exit Sub
Catch:
    LogError "Error", "CreateProject", Err.Description
    MsgBox "Project creation failed!", vbExclamation, "Error"
End Sub

'Private Sub AddItemsFromTemplates(strProjectName As String)
'    Dim DB As New OmlDatabase
'    Dim rst As ADODB.Recordset
'    Dim strItemName As String
'    Dim strItemFile As String
'    On Error GoTo Catch
'    With DB
'        .DataPath = gstrMasterDataPath & "\"
'        .DataFile = gstrMasterDataFile
'        .OpenMdb
'        ' FORMS
'        SQL_SELECT_ALL "Forms"
'        SQL_WHERE_Boolean "Active", True
'        SQL_AND_Text "TemplateName", strProjectType ' STANDARD
'        Set rst = .OpenRs(gstrSQL)
'        If .ErrorDesc <> "" Then
'            LogError "Error", "AddItemsFromTemplates/frmProjectNew", .ErrorDesc
'            .CloseRs rst
'            .CloseMdb
'            Exit Sub
'        End If
'        If Not (rst Is Nothing Or rst.EOF) Then
'            With rst
'                While Not .EOF
'                    strItemName = !FormName
'                    strItemFile = !FormFile
'                    If strItemName <> "" Then
'                        If strItemFile = "" Then strItemFile = strItemName & ".frm"
'                        If Not AddItem("Form", strItemName, gstrProjectPath, strItemFile) Then
'                            MsgBox "Database add new item (Form) failed", vbExclamation, "Error"
'                            Exit Sub
'                        End If
'                    End If
'                    .MoveNext
'                Wend
'            End With
'        End If
'        .CloseRs rst
'
'        ' MODULES
'        SQL_SELECT_ALL "Modules"
'        SQL_WHERE_Boolean "Active", True
'        SQL_AND_Text "TemplateName", strProjectType ' STANDARD
'        Set rst = .OpenRs(gstrSQL)
'        If .ErrorDesc <> "" Then
'            LogError "Error", "AddItemsFromTemplates/frmProjectNew", .ErrorDesc
'            .CloseRs rst
'            .CloseMdb
'            Exit Sub
'        End If
'        If Not (rst Is Nothing Or rst.EOF) Then
'            With rst
'                While Not .EOF
'                    strItemName = !ModuleName
'                    strItemFile = !ModuleFile
'                    If strItemName <> "" Then
'                        If strItemFile = "" Then strItemFile = strItemName & ".bas"
'                        If Not AddItem("Module", strItemName, gstrProjectPath, strItemFile) Then
'                            MsgBox "Database add new item (Module) failed", vbExclamation, "Error"
'                            Exit Sub
'                        End If
'                    End If
'                    .MoveNext
'                Wend
'            End With
'        End If
'        .CloseRs rst
'
'        ' CLASS
'        SQL_SELECT_ALL "Class"
'        SQL_WHERE_Boolean "Active", True
'        SQL_AND_Text "TemplateName", strProjectType ' STANDARD
'        Set rst = .OpenRs(gstrSQL)
'        If .ErrorDesc <> "" Then
'            LogError "Error", "AddItemsFromTemplates/frmProjectNew", .ErrorDesc
'            .CloseRs rst
'            .CloseMdb
'            Exit Sub
'        End If
'        If Not (rst Is Nothing Or rst.EOF) Then
'            With rst
'                While Not .EOF
'                    strItemName = !ClassName
'                    strItemFile = !ClassFile
'                    If strItemName <> "" Then
'                        If strItemFile = "" Then strItemFile = strItemName & ".cls"
'                        If Not AddItem("Class", strItemName, gstrProjectPath, strItemFile) Then
'                            MsgBox "Database add new item (Class) failed", vbExclamation, "Error"
'                            Exit Sub
'                        End If
'                    End If
'                    .MoveNext
'                Wend
'            End With
'        End If
'        .CloseRs rst
'        .CloseMdb
'    End With
'Exit Sub
'Catch:
'    LogError "Error", "AddItemsFromTemplates/frmProjectNew->", Err.Description
'End Sub

Private Sub GetItems(ByVal strTemplate As String, ByVal strItemType As String, ByRef pItem() As ProjectItem)
Try:
    On Error GoTo Catch
    Dim DB As New OmlDatabase
    Dim rst As ADODB.Recordset
    Dim strItem As String
    Dim strType As String
    Dim strFile As String
    With DB
        .DataPath = gstrMasterDataPath & "\"
        .DataFile = gstrMasterDataFile
        .OpenMdb
        Select Case strItemType
            Case "Forms"
                strItem = "Form"
                strType = ".frm"
            Case "Modules"
                strItem = "Module"
                strType = ".bas"
            Case "Class"
                strItem = "Class"
                strType = ".cls"
            Case "Project"
                strItem = "Project"
                strType = ".vbp"
            Case Else
                'strItem = ""
                'strType = ""
        End Select
        SQL_SELECT_ALL strItemType
        SQL_WHERE_Boolean "Active", True
        SQL_AND_Text "TemplateName", strTemplate
        Set rst = .OpenRs(gstrSQL)
        If .ErrorDesc <> "" Then
            LogError "Error", "GetItems/modScaffold", .ErrorDesc
            .CloseRs rst
            .CloseMdb
            Exit Sub
        End If
        If Not (rst Is Nothing Or rst.EOF) Then
            Dim i As Integer
            With rst
                ReDim pItem(.RecordCount - 1)
                While Not .EOF
                    If .Fields.Item(strItem & "Name").Value <> "" Then
                        pItem(i).Name = .Fields.Item(strItem & "Name").Value
                        strFile = .Fields.Item(strItem & "File").Value
                        SetDefaultValue strFile, .Fields.Item(strItem & "Name").Value & strType
                        pItem(i).File = strFile
                    End If
                    i = i + 1
                    .MoveNext
                Wend
            End With
        Else
            ReDim pItem(-1)
        End If
        .CloseRs rst
        .CloseMdb
    End With
Exit Sub
Catch:
    LogError "Error", "GetItems/modScaffold->", Err.Description
End Sub

'Public Sub AddItemsFromTemplates(strTemplate As String, strProjectName As String)
'    Dim DB As New OmlDatabase
'    Dim rst As ADODB.Recordset
'    Dim strItemName As String
'    Dim strItemFile As String
'Try:
'    On Error GoTo Catch
'    With DB
'        .DataPath = gstrMasterDataPath & "\"
'        .DataFile = gstrMasterDataFile
'        .OpenMdb
'        ' PROJECT
'
'        ' FORMS
'        SQL_SELECT_ALL "Forms"
'        SQL_WHERE_Boolean "Active", True
'        SQL_AND_Text "TemplateName", strTemplate
'        Set rst = .OpenRs(gstrSQL)
'        If .ErrorDesc <> "" Then
'            LogError "Error", "AddItemsFromTemplates/frmProjectNew", .ErrorDesc
'            .CloseRs rst
'            .CloseMdb
'            Exit Sub
'        End If
'        If Not (rst Is Nothing Or rst.EOF) Then
'            With rst
'                While Not .EOF
'                    strItemName = !FormName
'                    strItemFile = !FormFile
'                    If strItemName <> "" Then
'                        If strItemFile = "" Then strItemFile = strItemName & ".frm"
'                        If Not AddItem("Form", strItemName, gstrProjectPath, strItemFile) Then
'                            MsgBox "Database add new item (Form) failed", vbExclamation, "Error"
'                            Exit Sub
'                        End If
'                    End If
'                    .MoveNext
'                Wend
'            End With
'        End If
'        .CloseRs rst
'
'        ' MODULES
'        SQL_SELECT_ALL "Modules"
'        SQL_WHERE_Boolean "Active", True
'        SQL_AND_Text "TemplateName", strTemplate
'        Set rst = .OpenRs(gstrSQL)
'        If .ErrorDesc <> "" Then
'            LogError "Error", "AddItemsFromTemplates/frmProjectNew", .ErrorDesc
'            .CloseRs rst
'            .CloseMdb
'            Exit Sub
'        End If
'        If Not (rst Is Nothing Or rst.EOF) Then
'            With rst
'                While Not .EOF
'                    strItemName = !ModuleName
'                    strItemFile = !ModuleFile
'                    If strItemName <> "" Then
'                        If strItemFile = "" Then strItemFile = strItemName & ".bas"
'                        If Not AddItem("Module", strItemName, gstrProjectPath, strItemFile) Then
'                            MsgBox "Database add new item (Module) failed", vbExclamation, "Error"
'                            Exit Sub
'                        End If
'                    End If
'                    .MoveNext
'                Wend
'            End With
'        End If
'        .CloseRs rst
'
'        ' CLASS
'        SQL_SELECT_ALL "Class"
'        SQL_WHERE_Boolean "Active", True
'        SQL_AND_Text "TemplateName", strTemplate
'        Set rst = .OpenRs(gstrSQL)
'        If .ErrorDesc <> "" Then
'            LogError "Error", "AddItemsFromTemplates/frmProjectNew", .ErrorDesc
'            .CloseRs rst
'            .CloseMdb
'            Exit Sub
'        End If
'        If Not (rst Is Nothing Or rst.EOF) Then
'            With rst
'                While Not .EOF
'                    strItemName = !ClassName
'                    strItemFile = !ClassFile
'                    If strItemName <> "" Then
'                        If strItemFile = "" Then strItemFile = strItemName & ".cls"
'                        If Not AddItem("Class", strItemName, gstrProjectPath, strItemFile) Then
'                            MsgBox "Database add new item (Class) failed", vbExclamation, "Error"
'                            Exit Sub
'                        End If
'                    End If
'                    .MoveNext
'                Wend
'            End With
'        End If
'        .CloseRs rst
'
'        .CloseMdb
'    End With
'Exit Sub
'Catch:
'    LogError "Error", "AddItemsFromTemplates/frmProjectNew->", Err.Description
'End Sub

Public Function WriteVbp(pstrTemplate As String, _
                    pstrFilePath As String, _
                    pstrFileName As String, _
                    Optional pstrProjectName As String) As Boolean
    Dim DB As New OmlDatabase
    Dim rst As ADODB.Recordset
    Dim FF As Integer
    Dim strText As String
On Error GoTo Catch
    If pstrProjectName = "" Then pstrProjectName = pstrFileName
    With DB
        .DataPath = gstrMasterDataPath & "\"
        .DataFile = gstrMasterDataFile
        .OpenMdb
        SQL_SELECT_ALL "Code"
        SQL_WHERE_Boolean "Active", True
        If pstrTemplate = "ADMIN" Then
            SQL_AND_Text "FileName", "Admin.vbp"
        ElseIf pstrTemplate = "STANDARD" Then
            SQL_AND_Text "FileName", "Project1.vbp"
        Else
            SQL_AND_Text "FileName", pstrFileName
        End If
        SQL_AND_Text "MethodName", "WriteVbp"
        SQL_AND_Text "TemplateName", pstrTemplate
        SQL_ORDER_BY "ID" ' Make sure by Order
        Set rst = .OpenRs(gstrSQL)
        If .ErrorDesc <> "" Then
            LogError "Error", "WriteVbp/modScaffold", .ErrorDesc
            .CloseRs rst
            .CloseMdb
            WriteVbp = False
            Exit Function
        End If
        If Not (rst Is Nothing Or rst.EOF) Then
            FF = FreeFile
            Open pstrFilePath & "\" & pstrFileName For Output As #FF
            With rst
                While Not .EOF
                    strText = !CodeText
                    strText = Replace(strText, "{AppCompanyName}", App.CompanyName)
                    strText = Replace(strText, "{Parameter2}", pstrProjectName)
                    Print #FF, strText
                    .MoveNext
                Wend
            End With
            Close #FF
        End If
        .CloseRs rst
        .CloseMdb
    End With
    WriteVbp = True
    Exit Function
Catch:
    LogError "Error", "WriteVbp(" & pstrFileName & ")/modScaffold", Err.Description
    WriteVbp = False
End Function

' Write the file (overwrite existing file)
Public Function WriteFrm(pstrTemplate As String, pstrFilePath As String, pstrFormName As String) As Boolean
    Dim DB As New OmlDatabase
    Dim rst As ADODB.Recordset
    Dim FF As Integer
    Dim strText As String
    Dim strFileName As String
On Error GoTo Catch
    With DB
        strFileName = pstrFormName & ".frm"
        .DataPath = gstrMasterDataPath & "\"
        .DataFile = gstrMasterDataFile
        .OpenMdb
        SQL_SELECT_ALL "Code"
        SQL_WHERE_Boolean "Active", True
        SQL_AND_Text "FileName", strFileName
        SQL_AND_Text "MethodName", "WriteFrm"
        SQL_AND_Text "TemplateName", pstrTemplate
        ' Make sure by Order
        SQL_ORDER_BY "ID"
        'Debug.Print gstrSQL
        Set rst = .OpenRs(gstrSQL)
        If .ErrorDesc <> "" Then
            LogError "Error", "WriteFrm/modScaffold", .ErrorDesc
            .CloseRs rst
            .CloseMdb
            WriteFrm = False
            Exit Function
        End If
        If Not (rst Is Nothing Or rst.EOF) Then
            FF = FreeFile
            Open pstrFilePath & "\" & strFileName For Output As #FF
            With rst
                While Not .EOF
                    strText = !CodeText
                    strText = Replace(strText, "{Parameter2}", pstrFormName)
                    Print #FF, strText
                    .MoveNext
                Wend
            End With
            Close #FF
        End If
        .CloseRs rst
        .CloseMdb
    End With
    WriteFrm = True
    Exit Function
Catch:
    LogError "Error", "WriteFrm(" & pstrFormName & ")", Err.Description
    WriteFrm = False
End Function

Public Function WriteBas(pstrTemplate As String, pstrFilePath As String, pstrFileName As String) As Boolean
    Dim DB As New OmlDatabase
    Dim rst As ADODB.Recordset
    Dim FF As Integer
    Dim strText As String
On Error GoTo Catch
    With DB
        .DataPath = gstrMasterDataPath & "\"
        .DataFile = gstrMasterDataFile
        .OpenMdb
        SQL_SELECT_ALL "Code"
        SQL_WHERE_Boolean "Active", True
        SQL_AND_Text "FileName", pstrFileName
        SQL_AND_Text "MethodName", "WriteBas"
        SQL_AND_Text "TemplateName", pstrTemplate
        ' Make sure by Order
        SQL_ORDER_BY "ID"
        Set rst = .OpenRs(gstrSQL)
        If .ErrorDesc <> "" Then
            LogError "Error", "WriteBas/modProject", .ErrorDesc
            .CloseRs rst
            .CloseMdb
            WriteBas = False
            Exit Function
        End If
        If Not (rst Is Nothing Or rst.EOF) Then
            FF = FreeFile
            Open pstrFilePath & "\" & pstrFileName For Output As #FF
            With rst
                While Not .EOF
                    strText = !CodeText
                    Print #FF, strText
                    .MoveNext
                Wend
            End With
            Close #FF
        End If
        .CloseRs rst
        .CloseMdb
    End With
    WriteBas = True
    Exit Function
Catch:
    LogError "Error", "WriteBas(" & pstrFilePath & "\" & pstrFileName & ")", Err.Description
    WriteBas = False
End Function

Public Function WriteCls(pstrTemplate As String, pstrFilePath As String, pstrFileName As String) As Boolean
    Dim DB As New OmlDatabase
    Dim rst As ADODB.Recordset
    Dim FF As Integer
    Dim strText As String
On Error GoTo Catch
    With DB
        .DataPath = gstrMasterDataPath & "\"
        .DataFile = gstrMasterDataFile
        .OpenMdb
        SQL_SELECT_ALL "Code"
        SQL_WHERE_Boolean "Active", True
        SQL_AND_Text "FileName", pstrFileName
        SQL_AND_Text "MethodName", "WriteCls"
        SQL_AND_Text "TemplateName", pstrTemplate
        ' Make sure by Order
        SQL_ORDER_BY "ID"
        Set rst = .OpenRs(gstrSQL)
        If .ErrorDesc <> "" Then
            LogError "Error", "WriteCls/modProject", .ErrorDesc
            .CloseRs rst
            .CloseMdb
            WriteCls = False
            Exit Function
        End If
        If Not (rst Is Nothing Or rst.EOF) Then
            FF = FreeFile
            Open pstrFilePath & "\" & pstrFileName For Output As #FF
            With rst
                While Not .EOF
                    strText = !CodeText
                    Print #FF, strText
                    .MoveNext
                Wend
            End With
            Close #FF
        End If
        .CloseRs rst
        .CloseMdb
    End With
    WriteCls = True
    Exit Function
Catch:
    LogError "Error", "WriteCls(" & pstrFilePath & "\" & pstrFileName & ")", Err.Description
    WriteCls = False
End Function

' Continue write existing file - CRUD functions
'Public Sub WriteBas0001(pstrFileName As String)
'    Dim FF As Integer
'    Dim strText(6) As String
'On Error GoTo Catch
'    FF = FreeFile
'    strText(0) = "Public Function AddNew() As Boolean"
'    strText(1) = "    AddNew = True"
'    strText(2) = "End Function"
'    strText(3) = ""
'    strText(4) = "Public Function DeleteID() As Boolean"
'    strText(5) = "    DeleteID = True"
'    strText(6) = "End Function"
'    Open pstrFileName For Output As #FF
'        For i = 0 To 6
'            Print #FF, strText(i)
'        Next
'    Close #FF
'    Exit Sub
'Catch:
'    LogError "Error", "WriteBas0001(" & pstrFileName & ")", Err.Description
'End Sub
