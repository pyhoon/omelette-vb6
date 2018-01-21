Attribute VB_Name = "modDatabase"
' Version : 2.0
' Author: Poon Yip Hoon
' Created On   : Unknown
' Modified On  : 15 Oct 2017
' Descriptions : Commonly use database functions
' References:
' Microsoft ActiveX Data Objects 6.1 Library
' C:\Program Files (x86)\Common Files\System\ado\msado15.dll
' Can use older version such as 2.8
' Recommend to install VB6SP6 for latest version of libraries

' Dependencies:
' modVariable

Option Explicit
'Public gstrMasterDatabasePath As String
'Public gstrMasterDatabaseFile As String
'Public gstrMasterDatabasePassword As String
'Public gconMasterADODatabase As ADODB.Connection
'Public gstrDatabasePath As String
'Public gstrDatabaseFile As String
'Public gstrDatabasePassword As String
'Public gconADODatabase As ADODB.Connection

Public Sub ProjectOpenDB()
On Error GoTo CatchErr
    Set gconProject = New ADODB.Connection
    With gconProject
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & gstrProjectDataPath & "\" & gstrProjectDataFile
        .Properties("Jet OLEDB:Database Password") = gstrProjectDataPassword
        .Open
    End With
    Exit Sub
CatchErr:
    LogError "Error", "ProjectOpenDB", Err.Description
End Sub

Public Sub ProjectCloseDB()
On Error GoTo CatchErr
    If gconProject Is Nothing Then
        ' No action
    Else
        With gconProject
            If .State = adStateOpen Then
                .Close
            End If
        End With
        Set gconProject = Nothing
    End If
    Exit Sub
CatchErr:
    LogError "Error", "ProjectCloseDB", Err.Description
End Sub

Public Sub ProjectUnlockDB()
Dim rst As New ADODB.Recordset
On Error GoTo CatchErr
    ProjectOpenDB
    'gstrSQL = "SELECT * FROM NOTHING"
    gstrSQL = "SELECT * FROM UserData"
    rst.Open gstrSQL, gconProject, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Set UnlockDB = rst
    rst.Close
    Set rst = Nothing
    ProjectCloseDB
    Exit Sub
CatchErr:
    LogError "Error", "ProjectUnlockDB", Err.Description
End Sub
 
Public Sub ProjectQuerySQL(ByVal pstrSQL As String, Optional ByRef plngRecordsAffected As Long)
On Error GoTo CatchErr
    With gconProject
        .BeginTrans
        .Execute pstrSQL, plngRecordsAffected
        .CommitTrans
    End With
    Exit Sub
CatchErr:
    gconProject.RollbackTrans
    LogError "Error", "ProjectQuerySQL: " & pstrSQL, Err.Description
End Sub

Public Function ProjectOpenRS(ByVal pstrSQL As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CatchErr
    'rstSQL.Open pstrSQL, gconADODatabase, adOpenForwardOnly, adLockOptimistic, adCmdText
    rst.Open pstrSQL, gconProject, adOpenStatic, adLockPessimistic, adCmdText
    Set ProjectOpenRS = rst
    Exit Function
CatchErr:
    LogError "Error", "ProjectOpenRS: " & pstrSQL, Err.Description
End Function

Public Sub ProjectCloseRS(ByVal rst As ADODB.Recordset)
On Error GoTo CatchErr
    If rst Is Nothing Then
    Else
        If rst.State = adStateOpen Then
            rst.Close
        End If
        Set rst = Nothing
    End If
    Exit Sub
CatchErr:
    LogError "Error", "ProjectCloseRS", Err.Description
End Sub

Public Sub CloseRS(ByVal rst As ADODB.Recordset)
On Error GoTo CatchErr
    If rst Is Nothing Then
    Else
        If rst.State = adStateOpen Then
            rst.Close
        End If
        Set rst = Nothing
    End If
    Exit Sub
CatchErr:
    LogError "Error", "CloseRS", Err.Description
End Sub

Public Function ProjectOpenTable(ByVal pstrTable As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CatchErr
    rst.Open pstrTable, gconProject, adOpenDynamic, adLockPessimistic, adCmdTable
    Set ProjectOpenTable = rst
    Exit Function
CatchErr:
    LogError "Error", "ProjectOpenTable: " & pstrTable, Err.Description
End Function
 
Public Function ProjectOpenQuery(ByVal pstrSQL As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CatchErr
    rst.Open pstrSQL, gconProject, adOpenForwardOnly, adLockOptimistic, adCmdText 'adCmdStoredProc
    Set ProjectOpenQuery = rst
    Exit Function
CatchErr:
    LogError "Error", "ProjectOpenQuery: " & pstrSQL, Err.Description
End Function

Public Function ProjectOpenSQL(ByVal pconADO As ADODB.Connection, ByVal pstrSQL As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CatchErr
    rst.Open pstrSQL, gconProject, adOpenForwardOnly, adLockOptimistic, adCmdText
    Set ProjectOpenSQL = rst
    Exit Function
CatchErr:
    LogError "Error", "ProjectOpenSQL", Err.Description
End Function

'Public Function OpenProc(strProcedureName As String, paramName1 As String, paramType1 As String, paramValue1 As String, paramName2 As String, paramType2 As String, paramValue2 As String) As ADODB.Recordset
'Dim cmd As New ADODB.Command
'Dim rst As New ADODB.Recordset
'On Error GoTo CatchErr
'With cmd
'    .CommandText = strProcedureName
'    .CommandType = adCmdStoredProc
'    Set .ActiveConnection = gconADODatabase
'    If paramType1 = "Date" Then
'        .Parameters.Append .CreateParameter(paramName1, adDate, adParamInput, , paramValue1)
'    Else
'        .Parameters.Append .CreateParameter(paramName1, adVarChar, adParamInput, , paramValue1)
'    End If
'    If paramType2 = "Date" Then
'        .Parameters.Append .CreateParameter(paramName2, adDate, adParamInput, , paramValue2)
'    Else
'        .Parameters.Append .CreateParameter(paramName2, adVarChar, adParamInput, , paramValue2)
'    End If
'End With
'
'With rst
'    .CursorLocation = adUseClient
'    .Open cmd, , adOpenStatic, adLockReadOnly
'    'Set .ActiveConnection = Nothing
'End With
'    Set cmd = Nothing
'    Exit Function
'CatchErr:
'    LogError "Error", "OpenProc", Err.Description
'    'Debug.Print gstrSQL
'    Set cmd = Nothing
'End Function

'Public Function ExecuteSelectSQL(ByVal pconADO As ADODB.Connection, ByVal pstrSQL As String, Optional aiCursorType As Integer = adOpenDynamic) As ADODB.Recordset
'Dim rst As New ADODB.Recordset
'On Error GoTo CatchErr
'    If aiCursorType <> adOpenDynamic Then
'        rst.Open pstrSQL, pconADO, aiCursorType, adLockOptimistic
'    Else
'        'rst.Open pstrSQL, pconADO, adOpenForwardOnly, adLockOptimistic
'        rst.Open pstrSQL, pconADO, adOpenDynamic, adLockOptimistic
'    End If
'    Set ExecuteSelectSQL = rst
'    Exit Function
'CatchErr:
'    LogError "Error", "ExecuteSelectSQL", Err.Description
'End Function

'Public Function CompactDB(pstrOriginalFileName As String, pstrDestinationFileName As String) As Boolean
'    Dim oJetEngine As New JRO.JetEngine
'    Dim strSource As String
'    Dim strDestination As String
'On Error GoTo CatchErr
'    strSource = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'                "Data Source=" & pstrOriginalFileName & ";" & _
'                "Jet OLEDB:Database Password=" & gstrDatabasePassword & ";" & _
'                "Jet OLEDB:Engine Type=5;"
'
'    strDestination = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'              "Data Source=" & pstrDestinationFileName & ";" & _
'              "Jet OLEDB:Database Password=" & gstrDatabasePassword & ";" & _
'              "Jet OLEDB:Engine Type=5;"
'    'Set pconADO = Nothing
'    oJetEngine.CompactDatabase strSource, strDestination
'    Set oJetEngine = Nothing
'    CompactDB = True
'    Exit Function
'CatchErr:
'    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "CompactDB"
'    LogError "Error", "CompactDB", Err.Description
'    Set oJetEngine = Nothing
'    CompactDB = False
'End Function

Public Sub ItemOpenDB()
On Error GoTo CatchErr
    Set gconItem = New ADODB.Connection
    With gconItem
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & gstrProjectDataPath & gstrProjectItemsFile
        .Properties("Jet OLEDB:Database Password") = gstrProjectItemsPassword
        .Open
    End With
    Exit Sub
CatchErr:
    LogError "Error", "ItemOpenDB", Err.Description
End Sub

Public Sub ItemCloseDB()
On Error GoTo CatchErr
    If gconItem Is Nothing Then
        ' No action
    Else
        With gconItem
            If .State = adStateOpen Then
                .Close
            End If
        End With
        Set gconItem = Nothing
    End If
    Exit Sub
CatchErr:
    LogError "Error", "ItemCloseDB", Err.Description
End Sub

Public Function ItemOpenTable(ByVal pstrTable As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CatchErr
    rst.Open pstrTable, gconItem, adOpenDynamic, adLockPessimistic, adCmdTable
    Set ItemOpenTable = rst
    Exit Function
CatchErr:
    LogError "Error", "ItemOpenTable: " & pstrTable, Err.Description
End Function

Public Sub ItemUnlockDB()
Dim rst As New ADODB.Recordset
On Error GoTo CatchErr
    ItemOpenDB
    'gstrSQL = "SELECT * FROM NOTHING"
    gstrSQL = "SELECT * FROM Project"
    rst.Open gstrSQL, gconItem, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Set UnlockDB = rst
    rst.Close
    Set rst = Nothing
    ItemCloseDB
    Exit Sub
CatchErr:
    LogError "Error", "ItemUnlockDB", Err.Description
End Sub

Public Sub ItemQuerySQL(ByVal pstrSQL As String, Optional ByRef plngRecordsAffected As Long)
On Error GoTo CatchErr
    With gconItem
        .BeginTrans
        .Execute pstrSQL, plngRecordsAffected
        .CommitTrans
    End With
    Exit Sub
CatchErr:
    gconItem.RollbackTrans
    LogError "Error", "ItemQuerySQL: " & pstrSQL, Err.Description
End Sub

Public Sub MasterOpenDB()
On Error GoTo CatchErr
    Set gconMaster = New ADODB.Connection
    With gconMaster
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & gstrMasterDataPath & "\" & gstrMasterDataFile
        .Properties("Jet OLEDB:Database Password") = gstrMasterDataPassword
        .Open
    End With
    Exit Sub
CatchErr:
    LogError "Error", "MasterOpenDB", Err.Description
End Sub

Public Sub MasterCloseDB()
On Error GoTo CatchErr
    If gconMaster Is Nothing Then
        ' No action
    Else
        With gconMaster
            If .State = adStateOpen Then
                .Close
            End If
        End With
        Set gconMaster = Nothing
    End If
    Exit Sub
CatchErr:
    LogError "Error", "MasterCloseDB", Err.Description
End Sub

Public Function MasterOpenTable(ByVal pstrTable As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CatchErr
    rst.Open pstrTable, gconMaster, adOpenDynamic, adLockPessimistic, adCmdTable
    Set MasterOpenTable = rst
    Exit Function
CatchErr:
    LogError "Error", "MasterOpenTable: " & pstrTable, Err.Description
End Function

Public Function MasterOpenRS(ByVal pstrSQL As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CatchErr
    rst.Open pstrSQL, gconMaster, adOpenStatic, adLockPessimistic, adCmdText
    Set MasterOpenRS = rst
    Exit Function
CatchErr:
    LogError "Error", "MasterOpenRS: " & pstrSQL, Err.Description
End Function

Public Sub MasterUnlockDB()
Dim rst As New ADODB.Recordset
On Error GoTo CatchErr
    MasterOpenDB
    'gstrSQL = "SELECT * FROM NOTHING"
    gstrSQL = "SELECT * FROM UserData"
    rst.Open gstrSQL, gconMaster, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Set UnlockDB = rst
    rst.Close
    Set rst = Nothing
    MasterCloseDB
    Exit Sub
CatchErr:
    LogError "Error", "MasterUnlockDB", Err.Description
End Sub

Public Sub MasterQuerySQL(ByVal pstrSQL As String, Optional ByRef plngRecordsAffected As Long)
On Error GoTo CatchErr
    With gconMaster
        .BeginTrans
        .Execute pstrSQL, plngRecordsAffected
        .CommitTrans
    End With
    Exit Sub
CatchErr:
    gconMaster.RollbackTrans
    LogError "Error", "MasterQuerySQL: " & pstrSQL, Err.Description
End Sub

Public Sub SetRecord(ByVal txtOutput As TextBox, ByVal adoField As ADODB.Field, Optional ByVal blnTrim As Boolean = True)
    If adoField.Value <> "" Then
        If blnTrim = True Then
            txtOutput.Text = Trim(adoField)
        Else
            txtOutput.Text = adoField
        End If
    Else
        txtOutput.Text = ""
    End If
End Sub

Public Sub SetCheck(ByVal chkOutput As CheckBox, ByVal adoField As ADODB.Field)
    If adoField = True Then
        chkOutput.Value = vbChecked
    Else
        chkOutput.Value = vbUnchecked
    End If
End Sub
