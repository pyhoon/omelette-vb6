Attribute VB_Name = "modScaffold"
' Version : 0.2
' Author: Poon Yip Hoon
' Created On   : 18 Sep 2017
' Modified On  : 15 Oct 2017
' Descriptions : Scaffolding
' Dependencies:
' modDatabase
' modVariable

Option Explicit

' CLASS / MODEL
' Command: vb generate model Model_Name
'Public gstrModel As String

' CREATE - READ - UPDATE - DELETE
' Reference: https://msdn.microsoft.com/en-us/vba/access-vba/articles/create-and-delete-tables-and-indexes-using-access-sql

' ID Int, Name Text, [OPTIONAL] Type Text, Active Bit, CreatedDate SmallDateTime, CreatedBy Text, ModifiedDate SmallDateTime, ModifiedBy Text
Public Function CreateModel(ByVal strModel As String, _
                            Optional ByVal strPrefix As String = "", _
                            Optional ByVal blnAppendModel As Boolean = False, _
                            Optional ByVal blnActiveDefault As Boolean = True) As Boolean
    On Error GoTo Catch
    gstrSQL = "CREATE TABLE " & strPrefix & strModel
    gstrSQL = gstrSQL & " (["
    If blnAppendModel Then
        gstrSQL = gstrSQL & strModel
    End If
    'gstrSQL = gstrSQL & "ID] AUTOINCREMENT,"
    gstrSQL = gstrSQL & "ID] AUTOINCREMENT PRIMARY KEY,"   '"ID INTEGER NOT NULL,"
    gstrSQL = gstrSQL & " ["
    If blnAppendModel Then
        gstrSQL = gstrSQL & strModel
    End If
    gstrSQL = gstrSQL & "Name] TEXT(255) NOT NULL,"
    gstrSQL = gstrSQL & " [Active] BIT" ' YESNO
    If blnActiveDefault Then
        gstrSQL = gstrSQL & " DEFAULT -1" ' YES
    End If
    gstrSQL = gstrSQL & " NOT NULL,"
    'gstrSQL = gstrSQL & " [CreatedDate] DATETIME"
    gstrSQL = gstrSQL & " [CreatedDate] DATETIME DEFAULT NOW() NOT NULL"
    gstrSQL = gstrSQL & ")"
    'Debug.Print gstrSQL
    
    'Debug.Print gstrProjectName
    'Debug.Print gstrProjectDataPath
        
    ProjectOpenDB
    If Err.Description <> "" Then
        CreateModel = False
        Exit Function
    End If
    'ProjectQuerySQL gstrSQL
    Dim od As New clsDatabase
    od.ProjectQuerySQL gstrSQL
    If od.ErrorDesc <> "" Then
        CreateModel = False
        Exit Function
    End If
    ProjectCloseDB
    If Err.Description <> "" Then
        CreateModel = False
        Exit Function
    End If
'    MasterOpenDB
'    MasterQuerySQL gstrSQL
'    MasterCloseDB
    
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
    CreateModel = False
End Function

Public Function CreateItemsModel(ByVal strItemModel As String, _
                            Optional ByVal strItem As String = "", _
                            Optional ByVal blnActiveDefault As Boolean = True) As Boolean
    On Error GoTo Catch
    If strItem = "" Then strItem = strItemModel
    gstrSQL = "CREATE TABLE [" & strItemModel & "]"
    gstrSQL = gstrSQL & " ([ID] AUTOINCREMENT PRIMARY KEY,"   '"ID INTEGER NOT NULL,"
    gstrSQL = gstrSQL & " [" & strItem & "Name] TEXT(255),"
    gstrSQL = gstrSQL & " [" & strItem & "Path] TEXT(255),"
    gstrSQL = gstrSQL & " [" & strItem & "File] TEXT(255),"
    gstrSQL = gstrSQL & " [Active] BIT" ' YESNO
    If blnActiveDefault Then
        gstrSQL = gstrSQL & " DEFAULT -1" ' YES
    End If
    gstrSQL = gstrSQL & ","
    gstrSQL = gstrSQL & " [CreatedDate] DATETIME DEFAULT Now(),"
    gstrSQL = gstrSQL & " [CreatedAuthor] TEXT(255),"
    gstrSQL = gstrSQL & " [ModifiedDate] DATETIME,"
    gstrSQL = gstrSQL & " [ModifiedEditor] TEXT(255)"
    gstrSQL = gstrSQL & ")"
    'Debug.Print gstrSQL
    ItemOpenDB
    If Err.Description <> "" Then
        CreateItemsModel = False
        Exit Function
    End If
    'ProjectQuerySQL gstrSQL
    Dim od As New clsDatabase
    od.ItemQuerySQL gstrSQL
    If od.ErrorDesc <> "" Then
        CreateItemsModel = False
        Exit Function
    End If
    ItemCloseDB
    If Err.Description <> "" Then
        CreateItemsModel = False
        Exit Function
    End If
    CreateItemsModel = True
    Exit Function
Catch:
    CreateItemsModel = False
End Function

' 05 Dec 2017
Public Function CreateUserModel(Optional ByVal strUserModel As String = "", _
                            Optional ByVal strUserName As String = "", _
                            Optional ByVal strFirstName As String = "", _
                            Optional ByVal strLastName As String = "", _
                            Optional ByVal strUserPassword As String = "", _
                            Optional ByVal blnActiveDefault As Boolean = True) As Boolean
    On Error GoTo Catch
    If strUserModel = "" Then strUserModel = "User"
    If strUserName = "" Then strUserName = "UserName"
    If strFirstName = "" Then strFirstName = "FirstName"
    If strLastName = "" Then strLastName = "LastName"
    If strUserPassword = "" Then strUserPassword = "UserPassword"
    gstrSQL = "CREATE TABLE [" & strUserModel & "]"
    gstrSQL = gstrSQL & " ([ID] AUTOINCREMENT PRIMARY KEY,"   '"ID INTEGER NOT NULL,"
    gstrSQL = gstrSQL & " [" & strUserName & "] TEXT(255),"
    gstrSQL = gstrSQL & " [" & strFirstName & "] TEXT(255),"
    gstrSQL = gstrSQL & " [" & strLastName & "] TEXT(255),"
    gstrSQL = gstrSQL & " [" & strUserPassword & "] TEXT(255),"
    gstrSQL = gstrSQL & " [Active] BIT" ' YESNO
    If blnActiveDefault Then
        gstrSQL = gstrSQL & " DEFAULT -1" ' YES
    End If
    gstrSQL = gstrSQL & ","
    gstrSQL = gstrSQL & " [CreatedDate] DATETIME DEFAULT Now(),"
    gstrSQL = gstrSQL & " [CreatedAuthor] TEXT(255),"
    gstrSQL = gstrSQL & " [ModifiedDate] DATETIME,"
    gstrSQL = gstrSQL & " [ModifiedEditor] TEXT(255)"
    gstrSQL = gstrSQL & ")"
    'Debug.Print gstrSQL
    ProjectOpenDB
    If Err.Description <> "" Then
        CreateUserModel = False
        Exit Function
    End If
    'ProjectQuerySQL gstrSQL
    Dim od As New clsDatabase
    od.ProjectQuerySQL gstrSQL
    If od.ErrorDesc <> "" Then
        Err.Description = od.ErrorDesc
        CreateUserModel = False
        Exit Function
    End If
    ProjectCloseDB
    If Err.Description <> "" Then
        CreateUserModel = False
        Exit Function
    End If
    CreateUserModel = True
    Exit Function
Catch:
    Err.Description = od.ErrorDesc
    CreateUserModel = False
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

Public Function ProjectCreateDatabase() As Boolean
On Error GoTo Catch
    Dim cat As New ADOX.Catalog
    Dim strDBCon As String
    strDBCon = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strDBCon = strDBCon & "Data Source=" & gstrProjectDataPath & gstrProjectDataFile & ";"
    strDBCon = strDBCon & "Jet OLEDB:Database Password=" & gstrProjectDataPassword
    cat.Create strDBCon
    ProjectCreateDatabase = True
    Exit Function
Catch:
    ProjectCreateDatabase = False
    Set cat = Nothing
End Function

Public Function ItemsCreateDatabase() As Boolean
On Error GoTo Catch
    Dim cat As New ADOX.Catalog
    Dim strDBCon As String
    strDBCon = "Provider=Microsoft.Jet.OLEDB.4.0;"
    strDBCon = strDBCon & "Data Source=" & gstrProjectDataPath & gstrProjectItemsFile & ";"
    strDBCon = strDBCon & "Jet OLEDB:Database Password=" & gstrProjectItemsPassword
    cat.Create strDBCon
    ItemsCreateDatabase = True
    Exit Function
Catch:
    ItemsCreateDatabase = False
    Set cat = Nothing
End Function

Public Function ProjectAddNew(ByVal pstrProjectName As String, ByVal pstrProjectPath As String, ByVal pstrProjectFile As String) As Boolean
On Error GoTo Catch
    Dim rst As New ADODB.Recordset
    MasterOpenDB
    Set rst = MasterOpenTable("Project")
    With rst
        .AddNew
        !ProjectName = pstrProjectName
        !Projectpath = pstrProjectPath
        !ProjectFile = pstrProjectFile
        !CreatedDate = Now
        !CreatedAuthor = "Aeric" ' Change to your name
        .Update
    End With
    CloseRS rst
    MasterCloseDB
    ProjectAddNew = True
    Exit Function
Catch:
    ProjectAddNew = False
    'MasterCloseDB
End Function

Public Function ItemAddNew(ByVal pstrItemName As String, ByVal pstrItemPath As String, ByVal pstrItemFile As String, ByVal pstrItemType As String) As Boolean
On Error GoTo Catch
    Dim rst As New ADODB.Recordset
    ItemOpenDB
    If pstrItemType = "Project" Then
        Set rst = ItemOpenTable("Project")
    ElseIf pstrItemType = "Form" Then
        Set rst = ItemOpenTable("Forms")
    ElseIf pstrItemType = "Module" Then
        Set rst = ItemOpenTable("Modules")
    End If
    With rst
        .AddNew
        If pstrItemType = "Project" Then
            !ProjectName = pstrItemName
            !Projectpath = pstrItemPath
            !ProjectFile = pstrItemFile
        ElseIf pstrItemType = "Form" Then
            !FormName = pstrItemName
            !FormPath = pstrItemPath
            !FormFile = pstrItemFile
        ElseIf pstrItemType = "Module" Then
            !ModuleName = pstrItemName
            !ModulePath = pstrItemPath
            !ModuleFile = pstrItemFile
        End If
        !CreatedDate = Now
        !CreatedAuthor = "Aeric" ' Change to your name
        .Update
    End With
    CloseRS rst
    ItemCloseDB
    ItemAddNew = True
    Exit Function
Catch:
    LogError "Error", "ItemAddNew", Err.Description
    ItemAddNew = False
End Function
