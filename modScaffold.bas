Attribute VB_Name = "modScaffold"
' Version : 0.1
' Author: Poon Yip Hoon
' Modified On : 18 Sep 2017
' Descriptions : Scaffolding
' Dependencies:
' modDatabase

Option Explicit

' CLASS / MODEL
' Command: vb generate scaffold Model_Name
Public gstrModel As String

' CREATE - READ - UPDATE - DELETE
' Reference: https://msdn.microsoft.com/en-us/vba/access-vba/articles/create-and-delete-tables-and-indexes-using-access-sql

' ID Int, Name Text, [OPTIONAL] Type Text, Active Bit, CreatedDate SmallDateTime, CreatedBy Text, ModifiedDate SmallDateTime, ModifiedBy Text
Public Function CreateModel(strModel As String, Optional strPrefix As String = "", Optional blnAppendModel As Boolean, Optional blnActiveDefault As Boolean = True) As Boolean
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
    OpenDB
    QuerySQL gstrSQL
    CloseDB
    
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
End Function

Public Function CreateDatabase() As Boolean
On Error GoTo Catch
    Dim cat As New ADOX.Catalog
    Dim strDBCon As String
    'strDBPath = App.Path & "\Full.lic"
    strDBCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & gstrDatabasePath & "\" & gstrDatabaseFile
    strDBCon = strDBCon & ";Jet OLEDB:Database Password=" & gstrDatabasePassword
    cat.Create strDBCon
    CreateDatabase = True
    Exit Function
Catch:
      CreateDatabase = False
      'MsgBox Err.Description
      Set cat = Nothing
End Function
