Attribute VB_Name = "modScaffold"
' Version : 0.3
' Author: Poon Yip Hoon
' Created On   : 18 Sep 2017
' Modified On  : 22 Feb 2018
' Descriptions : Scaffolding
' Dependencies:
' modVariable
' OmlDatabase

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
    Dim DB As New OmlDatabase
    With DB
        .DataPath = gstrProjectDataPath & "\"
        .DataFile = gstrProjectDataFile
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
        SQL_COLUMN_TEXT "[CreatedBy]", 100, "Aeric"
        SQL_COLUMN_DATETIME "[ModifiedDate]"
        SQL_COLUMN_TEXT "[ModifiedBy]", 100, "", True, True, False
        SQL_Close_Bracket
        Debug.Print gstrSQL
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

Public Function CreateItemsModel(ByVal strItemModel As String, _
                            Optional ByVal strItem As String = "", _
                            Optional ByVal blnActiveDefault As Boolean = True) As Boolean
    On Error GoTo Catch
    If strItem = "" Then strItem = strItemModel
    Dim DB As New OmlDatabase
    With DB
        .DataPath = gstrProjectDataPath & "\"
        .DataFile = gstrProjectItemsFile
        .OpenMdb
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            CreateItemsModel = False
            Exit Function
        End If
        
        SQL_CREATE strItemModel
        SQL_COLUMN_ID
        SQL_COLUMN_TEXT "[" & strItem & "Name]"
        SQL_COLUMN_TEXT "[" & strItem & "Path]"
        If strItemModel = "Project" Then
            SQL_COLUMN_TEXT "[" & strItem & "Data]"
        End If
        SQL_COLUMN_TEXT "[" & strItem & "File]"
        SQL_COLUMN_YESNO "[Active]"
        SQL_COLUMN_DATETIME "[CreatedDate]", "NOW()"
        SQL_COLUMN_TEXT "[CreatedBy]", 255, "Aeric"
        SQL_COLUMN_DATETIME "[ModifiedDate]"
        SQL_COLUMN_TEXT "[ModifiedBy]", 255, "", True, True, False
        SQL_Close_Bracket
        
        Debug.Print gstrSQL
        .Execute gstrSQL
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            CreateItemsModel = False
            Exit Function
        End If
        .CloseMdb
    End With
    CreateItemsModel = True
    Exit Function
Catch:
    Err.Description = Err.Description & " " & DB.ErrorDesc
    CreateItemsModel = False
End Function

' 05 Dec 2017
' 22 Feb 2018
Public Function CreateUserModel(Optional ByVal strUserModel As String = "", _
                            Optional ByVal strUserName As String = "", _
                            Optional ByVal strFirstName As String = "", _
                            Optional ByVal strLastName As String = "", _
                            Optional ByVal strUserPassword As String = "", _
                            Optional ByVal strSalt As String = "", _
                            Optional ByVal strUserRole As String = "", _
                            Optional ByVal strActiveDefault As String = "") As Boolean
    On Error GoTo Catch
    If strUserModel = "" Then strUserModel = "[User]"
    If strUserName = "" Then strUserName = "[UserName]"
    If strFirstName = "" Then strFirstName = "[FirstName]"
    If strLastName = "" Then strLastName = "[LastName]"
    If strUserPassword = "" Then strUserPassword = "[UserPassword]"
    If strSalt = "" Then strSalt = "[Salt]"
    If strUserRole = "" Then strUserRole = "[UserRole]"
    Dim DB As New OmlDatabase
    With DB
        .DataPath = gstrProjectDataPath & "\"
        .DataFile = gstrProjectDataFile
        .OpenMdb
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            CreateUserModel = False
            Exit Function
        End If
        
        SQL_CREATE strUserModel
        SQL_COLUMN_ID
        SQL_COLUMN_TEXT strUserName
        SQL_COLUMN_TEXT strFirstName, 100
        SQL_COLUMN_TEXT strLastName, 100
        SQL_COLUMN_TEXT strUserPassword
        SQL_COLUMN_TEXT strSalt
        SQL_COLUMN_TEXT strUserRole
        SQL_COLUMN_YESNO "[Active]", strActiveDefault
        SQL_COLUMN_DATETIME "[CreatedDate]", "NOW()"
        SQL_COLUMN_TEXT "[CreatedBy]", 255, "Aeric"
        SQL_COLUMN_DATETIME "[ModifiedDate]"
        SQL_COLUMN_TEXT "[ModifiedBy]", 255, "", True, True, False
        SQL_Close_Bracket
        Debug.Print gstrSQL
        .Execute gstrSQL
        If .ErrorDesc <> "" Then
            Err.Description = Err.Description & " " & .ErrorDesc
            CreateUserModel = False
            Exit Function
        End If
        .CloseMdb
    End With
    CreateUserModel = True
    Exit Function
Catch:
    Err.Description = Err.Description & " " & DB.ErrorDesc
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
    strDBCon = strDBCon & "Data Source=" & gstrProjectDataPath & "\" & gstrProjectDataFile & ";"
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
    strDBCon = strDBCon & "Data Source=" & gstrProjectDataPath & "\" & gstrProjectItemsFile & ";"
    strDBCon = strDBCon & "Jet OLEDB:Database Password=" & gstrProjectItemsPassword
    cat.Create strDBCon
    ItemsCreateDatabase = True
    Exit Function
Catch:
    ItemsCreateDatabase = False
    Set cat = Nothing
End Function

Public Function ProjectAddNew(ByVal pstrProjectName As String, ByVal pstrProjectPath As String, ByVal pstrProjectData As String, ByVal pstrProjectFile As String) As Boolean
On Error GoTo Catch
    Dim DB As New OmlDatabase
    With DB
        .DataPath = gstrMasterDataPath & "\"
        .DataFile = gstrMasterDataFile
        .OpenMdb
        If .ErrorDesc <> "" Then
            ProjectAddNew = False
            Exit Function
        End If
        SQL_INSERT "Project"
        SQLText "ProjectName", True, False
        SQLText "ProjectPath"
        SQLText "ProjectData"
        SQLText "ProjectFile"
        SQLText "CreatedDate"
        SQLText "CreatedBy", False
        SQL_VALUES
        SQLData_Text pstrProjectName, True, False
        SQLData_Text pstrProjectPath
        SQLData_Text pstrProjectData
        SQLData_Text pstrProjectFile
        SQLData_DateTime Now
        SQLData_Text "Aeric", False  ' Change to your name
        SQL_Close_Bracket
        Debug.Print gstrSQL
        .Execute gstrSQL
        If .ErrorDesc <> "" Then
            .CloseMdb
            ProjectAddNew = False
            Exit Function
        End If
        .CloseMdb
    End With
    ProjectAddNew = True
    Exit Function
Catch:
    DB.CloseMdb
    ProjectAddNew = False
End Function

Public Function ItemAddNew(ByVal pstrItemName As String, ByVal pstrItemPath As String, ByVal pstrItemFile As String, ByVal pstrItemType As String, Optional ByVal pstrItemData As String = "Storage") As Boolean
On Error GoTo Catch
    Dim DB As New OmlDatabase
    With DB
        .DataPath = gstrProjectDataPath & "\"
        .DataFile = gstrProjectItemsFile
        .OpenMdb
        If .ErrorDesc <> "" Then
            ItemAddNew = False
            Exit Function
        End If
        If pstrItemType = "Project" Then
            SQL_INSERT "Project"
            SQLText "ProjectName", True, False
            SQLText "ProjectPath"
            SQLText "ProjectData"
            SQLText "ProjectFile"
            SQLText "CreatedDate"
            SQLText "CreatedBy", False
            SQL_VALUES
            SQLData_Text pstrItemName, True, False
            SQLData_Text pstrItemPath
            SQLData_Text pstrItemData
            SQLData_Text pstrItemFile
            SQLData_DateTime Now
            SQLData_Text "Aeric", False  ' Change to your name
            SQL_Close_Bracket
        ElseIf pstrItemType = "Form" Then
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
            SQLData_Text "Aeric", False  ' Change to your name
            SQL_Close_Bracket
        ElseIf pstrItemType = "Module" Then
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
            SQLData_Text "Aeric", False  ' Change to your name
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
            SQLData_Text "Aeric", False  ' Change to your name
            SQL_Close_Bracket
        End If
        Debug.Print gstrSQL
        .Execute gstrSQL
        If .ErrorDesc <> "" Then
            .CloseMdb
            ItemAddNew = False
            Exit Function
        End If
        .CloseMdb
    End With
    ItemAddNew = True
    Exit Function
Catch:
    LogError "Error", "ItemAddNew", Err.Description
    DB.CloseMdb
    ItemAddNew = False
End Function
