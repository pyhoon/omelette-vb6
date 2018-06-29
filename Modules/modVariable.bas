Attribute VB_Name = "modVariable"
' Global Declaration
' Public variables or constant naming connvention starts with g

'Const WIN_STATE_NORMAL = 0
'Const WIN_STATE_MINIMIZED = 1
'Const WIN_STATE_MAXIMIZED = 2

Public gstrProjectName As String
Public gstrProjectFile As String
Public gstrProjectFolder As String
Public gstrProjectData As String
Public gstrProjectPath As String
Public gstrProjectDataFile As String
Public gstrProjectDataPath As String
Public gstrProjectDataPassword As String
Public gstrProjectItemsFile As String
Public gstrProjectItemsPassword As String
Public gstrProjectClasses As String
Public gstrProjectModules As String

'Public gstrDatabasePath As String
'Public gstrDatabaseFile As String
'Public gstrDatabasePassword As String

'Public gstrMasterFolder As String ' App.Path
Public gstrMasterData As String
Public gstrMasterPath As String ' App.Path
Public gstrMasterDataFile As String
Public gstrMasterDataPath As String
Public gstrMasterDataPassword As String

Public gstrSQL As String

'Public gconADODatabase As ADODB.Connection
Public gconMaster As ADODB.Connection
Public gconProject As ADODB.Connection
Public gconItem As ADODB.Connection

Public gstrAppCompanyName As String
