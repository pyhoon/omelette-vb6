Attribute VB_Name = "modScaffold"
' Version : 0.5
' Author: Poon Yip Hoon
' Created On   : 18 Sep 2017
' Modified On  : 24 Jan 2019
' Descriptions : Scaffolding
' Dependencies:
' modProject
' modVariable
' OmlDatabase
' OmlSQLBuilder

Option Explicit

' CLASS / MODEL
' Command: vb generate model Model_Name
'Public gstrModel As String

' CREATE - READ - UPDATE - DELETE
' Reference: https://msdn.microsoft.com/en-us/vba/access-vba/articles/create-and-delete-tables-and-indexes-using-access-sql

' ID Int, Name Text, [OPTIONAL] Type Text, Active Bit, CreatedDate SmallDateTime, CreatedBy Text, ModifiedDate SmallDateTime, ModifiedBy Text

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
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
On Error GoTo Catch
Try:
    SetDefaultValue DataPath, gstrProjectDataPath
    SetDefaultValue DataFile, gstrProjectDataFile
    DB.DataPath = DataPath & "\"
    DB.DataFile = DataFile
    DB.OpenMdb
    If DB.ErrorDesc <> "" Then
        Err.Description = Err.Description & " " & DB.ErrorDesc
        LogError "Error", "CreateModel/modScaffold", DB.ErrorDesc
        CreateModel = False
        Exit Function
    End If
    SQ.CREATE strModel, strPrefix
    If blnAppendModel Then
        SQ.COLUMN_ID strModel & "ID"
        SQ.COLUMN_TEXT "[" & strModel & "Name]" ' First Column
        'SQ.COLUMN_TEXT "[" & strModel & "Type]"  ' Second Column
    Else
        SQ.COLUMN_ID
        SQ.COLUMN_TEXT "[Name]"
        'SQ.COLUMN_TEXT "[Type]"
    End If
    SQ.COLUMN_YESNO "[Active]"
    SQ.COLUMN_DATETIME "[CreatedDate]", "NOW()"
    SQ.COLUMN_TEXT "[CreatedBy]", 100, "Omelette"  ' Change to your name
    SQ.COLUMN_DATETIME "[ModifiedDate]"
    SQ.COLUMN_TEXT "[ModifiedBy]", 100, "", -1, -1, 0
    SQ.Close_Bracket
    Debug.Print SQ.Text
    DB.Execute SQ.Text
    If DB.ErrorDesc <> "" Then
        Err.Description = Err.Description & " " & DB.ErrorDesc
        LogError "Error", "CreateModel/modScaffold", DB.ErrorDesc
        CreateModel = False
        Exit Function
    End If
    DB.CloseMdb
    
    'SQ.SQL "CREATE INDEX idx" & strModel & "ID", 0, 0
    'SQ.SQL "ON " & strPrefix & strModel, 0
    'SQ.SOB " ", 0
    'If blnAppendModel Then SQ.SQL strModel
    'SQ.SQL "ID"
    'SQ.SCB ""
    'SQ.SQL "WITH PRIMARY", 0
    'DB.OpenMdb
    'DB.Execute SQ.Text
    'DB.CloseMdb
    
    'DB.OpenMdb
    'SQ.ALTER_TABLE strPrefix & strModel
    'SQ.SQL "ALTER COLUMN", 0
    'SQ.SQL "[CreatedDate] DATETIME DEFAULT NOW() NOT NULL", 0
    'DB.Execute SQ.Text
    'DB.CloseMdb
    
    CreateModel = True
    Exit Function
Catch:
    Err.Description = Err.Description & " " & DB.ErrorDesc
    LogError "Error", "CreateModel/modScaffold", Err.Description
    CreateModel = False
End Function

Public Function CreateItems(ByVal strItemModel As String, _
                            Optional ByVal strItem As String = "", _
                            Optional ByVal blnActiveDefault As Boolean = True, _
                            Optional ByVal DataPath As String = "", _
                            Optional ByVal DataFile As String = "") As Boolean
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
On Error GoTo Catch
Try:
    If strItem = "" Then strItem = strItemModel
    SetDefaultValue DataPath, gstrProjectDataPath
    SetDefaultValue DataFile, gstrProjectItemsFile
    DB.DataPath = DataPath & "\"
    DB.DataFile = DataFile
    DB.OpenMdb
    If DB.ErrorDesc <> "" Then
        LogError "Error", "CreateItems/modScaffold", DB.ErrorDesc
        Err.Description = Err.Description & " " & DB.ErrorDesc
        CreateItems = False
        Exit Function
    End If
    SQ.CREATE strItemModel
    SQ.COLUMN_ID
    SQ.COLUMN_TEXT "[" & strItem & "Name]"
    SQ.COLUMN_TEXT "[" & strItem & "Path]"
    SQ.COLUMN_TEXT "[" & strItem & "File]"
    If strItemModel = "Project" Then SQ.COLUMN_TEXT "[" & strItem & "Data]" ' Default Data Folder i.e. Storage
    SQ.COLUMN_YESNO "[Active]"
    SQ.COLUMN_DATETIME "[CreatedDate]", "NOW()"
    SQ.COLUMN_TEXT "[CreatedBy]", 255, "Omelette"  ' Change to your name
    SQ.COLUMN_DATETIME "[ModifiedDate]"
    SQ.COLUMN_TEXT "[ModifiedBy]", 255, "", True, True, False
    SQ.Close_Bracket
    Debug.Print SQ.Text
    DB.Execute SQ.Text
    If DB.ErrorDesc <> "" Then
        LogError "Error", "CreateItems/modScaffold", DB.ErrorDesc
        Err.Description = Err.Description & " " & DB.ErrorDesc
        CreateItems = False
        Exit Function
    End If
    DB.CloseMdb
    CreateItems = True
    Exit Function
Catch:
    LogError "Error", "CreateItems/modScaffold", Err.Description
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
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
On Error GoTo Catch
Try:
    SetDefaultValue strModel, "Users"
    SetDefaultValue strModelID, "ID"
    SetDefaultValue strUserID, "UserID"
    SetDefaultValue strUserName, "UserName"
    SetDefaultValue strFirstName, "FirstName"
    SetDefaultValue strLastName, "LastName"
    SetDefaultValue strUserPassword, "UserPassword"
    SetDefaultValue strSalt, "Salt"
    SetDefaultValue strUserRole, "UserRole"
    SetDefaultValue DataPath, gstrProjectDataPath
    SetDefaultValue DataFile, gstrProjectDataFile
    DB.DataPath = DataPath & "\"
    DB.DataFile = DataFile
    DB.OpenMdb
    If DB.ErrorDesc <> "" Then
        Err.Description = Err.Description & " " & DB.ErrorDesc
        CreateUsers = False
        Exit Function
    End If
        
    SQ.CREATE strModel
    SQ.COLUMN_ID strModelID
    SQ.COLUMN_TEXT strUserID
    SQ.COLUMN_TEXT strUserName
    SQ.COLUMN_TEXT strFirstName, 100
    SQ.COLUMN_TEXT strLastName, 100
    SQ.COLUMN_TEXT strUserPassword
    SQ.COLUMN_TEXT strSalt
    SQ.COLUMN_TEXT strUserRole
    SQ.COLUMN_YESNO "[Active]", strActiveDefault
    SQ.COLUMN_DATETIME "[CreatedDate]", "NOW()"
    SQ.COLUMN_TEXT "[CreatedBy]", 255, "Omelette"  ' Change to your name
    SQ.COLUMN_DATETIME "[ModifiedDate]"
    SQ.COLUMN_TEXT "[ModifiedBy]", 255, "", True, True, False
    SQ.Close_Bracket
    Debug.Print SQ.Text
    DB.Execute SQ.Text
    If DB.ErrorDesc <> "" Then
        Err.Description = Err.Description & " " & DB.ErrorDesc
        CreateUsers = False
        Exit Function
    End If
    DB.CloseMdb
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
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
On Error GoTo Catch
Try:
    DB.DataPath = gstrProjectDataPath & "\"
    DB.DataFile = gstrProjectDataFile
    DB.OpenMdb
    If DB.ErrorDesc <> "" Then
        Err.Description = Err.Description & " " & DB.ErrorDesc
        FindTable = False
        Exit Function
    End If
    SQ.SELECT_ID strTable
    Debug.Print SQ.Text
    DB.Execute SQ.Text
    If DB.ErrorDesc <> "" Then
        Err.Description = Err.Description & " " & DB.ErrorDesc
        FindTable = False
        Exit Function
    End If
    DB.CloseMdb
    FindTable = True
    Exit Function
Catch:
    Err.Description = Err.Description & " " & DB.ErrorDesc
    FindTable = False
End Function

Public Function DropTable(ByVal strTable As String) As Boolean
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
On Error GoTo Catch
Try:
    DB.DataPath = gstrProjectDataPath & "\"
    DB.DataFile = gstrProjectDataFile
    DB.OpenMdb
    If DB.ErrorDesc <> "" Then
        Err.Description = Err.Description & " " & DB.ErrorDesc
        DropTable = False
        Exit Function
    End If
    
    SQ.DROP strTable
    Debug.Print SQ.Text
    DB.Execute SQ.Text
    If DB.ErrorDesc <> "" Then
        Err.Description = Err.Description & " " & DB.ErrorDesc
        DropTable = False
        Exit Function
    End If
    DB.CloseMdb
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
    cat.CREATE strDBCon
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
    cat.CREATE strDBCon
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
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
On Error GoTo Catch
Try:
    DB.DataPath = gstrMasterDataPath & "\"
    DB.DataFile = gstrMasterDataFile
    DB.OpenMdb
    If DB.ErrorDesc <> "" Then
        AddProject = False
        Exit Function
    End If
    SQ.INSERT "Project"
    SQ.SOB "ProjectName"
    SQ.SQL "ProjectPath"
    SQ.SQL "ProjectFile"
    SQ.SQL "ProjectData"
    SQ.SQL "CreatedDate"
    SQ.SCB "CreatedBy"
    SQ.VALUES
    SQ.VOB pstrProjectName
    SQ.VAL pstrProjectPath
    SQ.VAL pstrProjectFile
    SQ.VAL pstrProjectData
    SQ.VDT Now
    SQ.VCB "Omelette" ' Change to your name
    Debug.Print SQ.Text
    DB.Execute SQ.Text
    If DB.ErrorDesc <> "" Then
        LogError "Error", "AddProject/modScaffold", DB.ErrorDesc
        DB.CloseMdb
        AddProject = False
        Exit Function
    End If
    DB.CloseMdb
    AddProject = True
    Exit Function
Catch:
    LogError "Error", "AddProject/modScaffold", Err.Description
    DB.CloseMdb
    AddProject = False
End Function

Public Function AddItem(ByVal pstrItemType As String, _
                            ByVal pstrItemName As String, _
                            ByVal pstrItemPath As String, _
                            ByVal pstrItemFile As String, _
                            Optional ByVal pstrItemData As String = "Storage", _
                            Optional ByVal pstrProjectName As String = "") As Boolean
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
Dim strDataPath As String
Dim strItemFile As String
On Error GoTo Catch
Try:
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
    DB.DataPath = strDataPath & "\"
    DB.DataFile = strItemFile
    DB.OpenMdb
    If DB.ErrorDesc <> "" Then
        AddItem = False
        Exit Function
    End If
    If pstrItemType = "Project" Then
        SQ.INSERT "Project"
        SQ.SOB "ProjectName"
        SQ.SQL "ProjectPath"
        SQ.SQL "ProjectFile"
        SQ.SQL "ProjectData"
        SQ.SQL "CreatedDate"
        SQ.SCB "CreatedBy"
        SQ.VALUES
        SQ.VOB pstrItemName
        SQ.VAL pstrItemPath
        SQ.VAL pstrItemFile
        SQ.VAL pstrItemData
        SQ.VDT Now
        SQ.VCB "Omelette"  ' Change to your name
    ElseIf pstrItemType = "Forms" Then
        SQ.INSERT "Forms"
        SQ.SOB "FormName"
        SQ.SQL "FormPath"
        SQ.SQL "FormFile"
        SQ.SQL "CreatedDate"
        SQ.SCB "CreatedBy"
        SQ.VALUES
        SQ.VOB pstrItemName
        SQ.VAL pstrItemPath
        SQ.VAL pstrItemFile
        SQ.VDT Now
        SQ.VCB "Omelette"  ' Change to your name
    ElseIf pstrItemType = "Modules" Then
        SQ.INSERT "Modules"
        SQ.SOB "ModuleName"
        SQ.SQL "ModulePath"
        SQ.SQL "ModuleFile"
        SQ.SQL "CreatedDate"
        SQ.SCB "CreatedBy"
        SQ.VALUES
        SQ.VOB pstrItemName
        SQ.VAL pstrItemPath
        SQ.VAL pstrItemFile
        SQ.VDT Now
        SQ.VCB "Omelette"  ' Change to your name
    ElseIf pstrItemType = "Class" Then ' Class Module
        SQ.INSERT "Class"
        SQ.SOB "ClassName"
        SQ.SQL "ClassPath"
        SQ.SQL "ClassFile"
        SQ.SQL "CreatedDate"
        SQ.SCB "CreatedBy"
        SQ.VALUES
        SQ.VOB pstrItemName
        SQ.VAL pstrItemPath
        SQ.VAL pstrItemFile
        SQ.VDT Now
        SQ.VCB "Omelette"  ' Change to your name
    ElseIf pstrItemType = "Data" Then
        SQ.INSERT "Data"
        SQ.SOB "DataName"
        SQ.SQL "DataPath"
        SQ.SQL "DataFile"
        SQ.SQL "CreatedDate"
        SQ.SCB "CreatedBy"
        SQ.VALUES
        SQ.VOB pstrItemName
        SQ.VAL pstrItemPath
        SQ.VAL pstrItemFile
        SQ.VDT Now
        SQ.VCB "Omelette"  ' Change to your name
    End If
    'Debug.Print SQ.Text
    DB.Execute SQ.Text
    If DB.ErrorDesc <> "" Then
        LogError "Error", "AddItem/modScaffold", DB.ErrorDesc '& vbCrLf & SQ.Text
        DB.CloseMdb
        AddItem = False
        Exit Function
    End If
    DB.CloseMdb
    AddItem = True
    Exit Function
Catch:
    LogError "Error", "AddItem/modScaffold", Err.Description
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
Dim pForm() As ProjectItem
Dim pModule() As ProjectItem
Dim pClass() As ProjectItem
Dim i As Integer
On Error GoTo Catch
Try:
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

Private Sub GetItems(ByVal strTemplate As String, ByVal strItemType As String, ByRef pItem() As ProjectItem)
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
Dim rst As ADODB.Recordset
Dim strItem As String
Dim strType As String
Dim strFile As String
Dim i As Integer
On Error GoTo Catch
Try:
    DB.DataPath = gstrMasterDataPath & "\"
    DB.DataFile = gstrMasterDataFile
    DB.OpenMdb
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
    SQ.SELECT_ALL strItemType
    SQ.WHERE_Boolean "Active", True
    SQ.AND_Text "TemplateName", strTemplate
    Set rst = DB.OpenRs(SQ.Text)
    If DB.ErrorDesc <> "" Then
        LogError "Error", "GetItems/modScaffold", DB.ErrorDesc
        DB.CloseRs rst
        DB.CloseMdb
        Exit Sub
    End If
    If Not (rst Is Nothing Or rst.EOF) Then
        ReDim pItem(rst.RecordCount - 1)
        While Not rst.EOF
            If rst.Fields.Item(strItem & "Name").Value <> "" Then
                pItem(i).Name = rst.Fields.Item(strItem & "Name").Value
                strFile = rst.Fields.Item(strItem & "File").Value
                SetDefaultValue strFile, rst.Fields.Item(strItem & "Name").Value & strType
                pItem(i).File = strFile
            End If
            i = i + 1
            rst.MoveNext
        Wend
    Else
        ReDim pItem(-1)
    End If
    DB.CloseRs rst
    DB.CloseMdb
Exit Sub
Catch:
    LogError "Error", "GetItems/modScaffold->", Err.Description
End Sub

Public Function WriteVbp(pstrTemplate As String, _
                    pstrFilePath As String, _
                    pstrFileName As String, _
                    Optional pstrProjectName As String) As Boolean
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
Dim rst As ADODB.Recordset
Dim FF As Integer
Dim strText As String
On Error GoTo Catch
Try:
    If pstrProjectName = "" Then pstrProjectName = pstrFileName
    DB.DataPath = gstrMasterDataPath & "\"
    DB.DataFile = gstrMasterDataFile
    DB.OpenMdb
    SQ.SELECT_ALL "Code"
    SQ.WHERE_Boolean "Active", True
    If pstrTemplate = "ADMIN" Then
        SQ.AND_Text "FileName", "Admin.vbp"
    ElseIf pstrTemplate = "STANDARD" Then
        SQ.AND_Text "FileName", "Project1.vbp"
    Else
        SQ.AND_Text "FileName", pstrFileName
    End If
    SQ.AND_Text "MethodName", "WriteVbp"
    SQ.AND_Text "TemplateName", pstrTemplate
    SQ.ORDER_BY "ID" ' Make sure by Order
    Set rst = DB.OpenRs(SQ.Text)
    If DB.ErrorDesc <> "" Then
        LogError "Error", "WriteVbp/modScaffold", DB.ErrorDesc
        DB.CloseRs rst
        DB.CloseMdb
        WriteVbp = False
        Exit Function
    End If
    If Not (rst Is Nothing Or rst.EOF) Then
        FF = FreeFile
        Open pstrFilePath & "\" & pstrFileName For Output As #FF
        While Not rst.EOF
            strText = rst!CodeText
            strText = Replace(strText, "{AppCompanyName}", App.CompanyName)
            strText = Replace(strText, "{Parameter2}", pstrProjectName)
            Print #FF, strText
            rst.MoveNext
        Wend
        Close #FF
    End If
    DB.CloseRs rst
    DB.CloseMdb
    WriteVbp = True
    Exit Function
Catch:
    LogError "Error", "WriteVbp(" & pstrFileName & ")/modScaffold", Err.Description
    WriteVbp = False
End Function

' Write the file (overwrite existing file)
Public Function WriteFrm(pstrTemplate As String, pstrFilePath As String, pstrFormName As String) As Boolean
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
Dim rst As ADODB.Recordset
Dim FF As Integer
Dim strText As String
Dim strFileName As String
On Error GoTo Catch
Try:
    strFileName = pstrFormName & ".frm"
    DB.DataPath = gstrMasterDataPath & "\"
    DB.DataFile = gstrMasterDataFile
    DB.OpenMdb
    SQ.SELECT_ALL "Code"
    SQ.WHERE_Boolean "Active", True
    SQ.AND_Text "FileName", strFileName
    SQ.AND_Text "MethodName", "WriteFrm"
    SQ.AND_Text "TemplateName", pstrTemplate
    ' Make sure by Order
    SQ.ORDER_BY "ID"
    Debug.Print SQ.Text
    Set rst = DB.OpenRs(SQ.Text)
    If DB.ErrorDesc <> "" Then
        LogError "Error", "WriteFrm/modScaffold", DB.ErrorDesc
        DB.CloseRs rst
        DB.CloseMdb
        WriteFrm = False
        Exit Function
    End If
    If Not (rst Is Nothing Or rst.EOF) Then
        FF = FreeFile
        Open pstrFilePath & "\" & strFileName For Output As #FF
        While Not rst.EOF
            strText = rst!CodeText
            strText = Replace(strText, "{Parameter2}", pstrFormName)
            Print #FF, strText
            rst.MoveNext
        Wend
        Close #FF
    End If
    DB.CloseRs rst
    DB.CloseMdb
    WriteFrm = True
    Exit Function
Catch:
    LogError "Error", "WriteFrm(" & pstrFormName & ")", Err.Description
    WriteFrm = False
End Function

Public Function WriteBas(pstrTemplate As String, pstrFilePath As String, pstrFileName As String) As Boolean
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
Dim rst As ADODB.Recordset
Dim FF As Integer
Dim strText As String
On Error GoTo Catch
Try:
    DB.DataPath = gstrMasterDataPath & "\"
    DB.DataFile = gstrMasterDataFile
    DB.OpenMdb
    SQ.SELECT_ALL "Code"
    SQ.WHERE_Boolean "Active", True
    SQ.AND_Text "FileName", pstrFileName
    SQ.AND_Text "MethodName", "WriteBas"
    SQ.AND_Text "TemplateName", pstrTemplate
    ' Make sure by Order
    SQ.ORDER_BY "ID"
    Set rst = DB.OpenRs(SQ.Text)
    If DB.ErrorDesc <> "" Then
        LogError "Error", "WriteBas/modScaffold", DB.ErrorDesc
        DB.CloseRs rst
        DB.CloseMdb
        WriteBas = False
        Exit Function
    End If
    If Not (rst Is Nothing Or rst.EOF) Then
        FF = FreeFile
        Open pstrFilePath & "\" & pstrFileName For Output As #FF
        While Not rst.EOF
            strText = rst!CodeText
            Print #FF, strText
            rst.MoveNext
        Wend
        Close #FF
    End If
    DB.CloseRs rst
    DB.CloseMdb
    WriteBas = True
    Exit Function
Catch:
    LogError "Error", "WriteBas(" & pstrFilePath & "\" & pstrFileName & ")", Err.Description
    WriteBas = False
End Function

Public Function WriteCls(pstrTemplate As String, pstrFilePath As String, pstrFileName As String) As Boolean
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
Dim rst As ADODB.Recordset
Dim FF As Integer
Dim strText As String
On Error GoTo Catch
    DB.DataPath = gstrMasterDataPath & "\"
    DB.DataFile = gstrMasterDataFile
    DB.OpenMdb
    SQ.SELECT_ALL "Code"
    SQ.WHERE_Boolean "Active", True
    SQ.AND_Text "FileName", pstrFileName
    SQ.AND_Text "MethodName", "WriteCls"
    SQ.AND_Text "TemplateName", pstrTemplate
    ' Make sure by Order
    SQ.ORDER_BY "ID"
    Set rst = DB.OpenRs(SQ.Text)
    If DB.ErrorDesc <> "" Then
        LogError "Error", "WriteCls/modScaffold", DB.ErrorDesc
        DB.CloseRs rst
        DB.CloseMdb
        WriteCls = False
        Exit Function
    End If
    If Not (rst Is Nothing Or rst.EOF) Then
        FF = FreeFile
        Open pstrFilePath & "\" & pstrFileName For Output As #FF
        While Not rst.EOF
            strText = rst!CodeText
            Print #FF, strText
            rst.MoveNext
        Wend
        Close #FF
    End If
    DB.CloseRs rst
    DB.CloseMdb
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
