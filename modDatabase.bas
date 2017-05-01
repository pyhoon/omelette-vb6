Attribute VB_Name = "modDatabase"
' Version : Unknown
' Author: Poon Yip Hoon
' Modified On : Unknown
' Descriptions : Commonly use database functions
' References:
' Microsoft ActiveX Data Objects 6.1 Library
' C:\Program Files (x86)\Common Files\System\ado\msado15.dll
' Can use older version such as 2.8
' Recommend to install VB6SP6 for latest version of libraries

Option Explicit
Public gstrDatabasePath As String
Public gstrDatabaseFile As String
Public gstrDatabasePassword As String
Public ACN As ADODB.Connection

Public Sub OpenDB()
On Error GoTo CheckErr
    Set ACN = New ADODB.Connection
    ACN.Provider = "Microsoft.Jet.OLEDB.4.0"
    ACN.ConnectionString = "Data Source=" & gstrDatabasePath & "\" & gstrDatabaseFile
    ACN.Properties("Jet OLEDB:Database Password") = gstrDatabasePassword
    ACN.Open
    Exit Sub
CheckErr:
    LogError "Error", "OpenDB", Err.Description
End Sub

Public Sub CloseDB()
On Error GoTo CheckErr
    If ACN Is Nothing Then
    Else
        If ACN.State = adStateOpen Then
            ACN.Close
        End If
        Set ACN = Nothing
    End If
    Exit Sub
CheckErr:
    LogError "Error", "CloseDB", Err.Description
End Sub

Public Function ExecuteSelectSQL(pstrSQL As String, Optional aiCursorType As Integer = adOpenDynamic) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    If aiCursorType <> adOpenDynamic Then
        rst.Open pstrSQL, ACN, aiCursorType, adLockOptimistic
    Else
        rst.Open pstrSQL, ACN, adOpenForwardOnly, adLockOptimistic
    End If
    Set ExecuteSelectSQL = rst
    Exit Function
CheckErr:
    LogError "Error", "ExecuteSelectSQL", Err.Description
End Function

Public Sub UnlockDB(Con As ADODB.Connection)
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    OpenDB
    'gstrSQL = "SELECT * FROM NOTHING"
    gstrSQL = "SELECT * FROM UserData"
    rst.Open gstrSQL, Con, adOpenForwardOnly, adLockOptimistic, adCmdText
    'Set UnlockDB = rst
    rst.Close
    Set rst = Nothing
    CloseDB
    Exit Sub
CheckErr:
    LogError "Error", "UnlockDB", Err.Description
End Sub
 
Public Function OpenSQL(pstrSQL As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    rst.Open pstrSQL, ACN, adOpenForwardOnly, adLockOptimistic, adCmdText
    Set OpenSQL = rst
    Exit Function
CheckErr:
    LogError "Error", "OpenSQL", Err.Description
End Function

Public Function OpenQuery(pstrSQL As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    rst.Open pstrSQL, ACN, adOpenForwardOnly, adLockOptimistic, adCmdText 'adCmdStoredProc
    Set OpenQuery = rst
    Exit Function
CheckErr:
    LogError "Error", "OpenQuery", Err.Description
End Function

'Public Function OpenProc(strProcedureName As String, paramName1 As String, paramType1 As String, paramValue1 As String, paramName2 As String, paramType2 As String, paramValue2 As String) As ADODB.Recordset
'Dim cmd As New ADODB.Command
'Dim rst As New ADODB.Recordset
'On Error GoTo CheckErr
'With cmd
'    .CommandText = strProcedureName
'    .CommandType = adCmdStoredProc
'    Set .ActiveConnection = ACN
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
'CheckErr:
'    LogError "Error", "OpenProc", Err.Description
'    'Debug.Print gstrSQL
'    Set cmd = Nothing
'End Function

Public Function OpenRS(pstrSQL As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    'rstSQL.Open pstrSQL, ACN, adOpenForwardOnly, adLockOptimistic, adCmdText
    rst.Open pstrSQL, ACN, adOpenStatic, adLockPessimistic, adCmdText
    Set OpenRS = rst
    Exit Function
CheckErr:
    LogError "Error", "OpenRS: " & pstrSQL, Err.Description
End Function

Public Sub CloseRS(rst As ADODB.Recordset)
On Error GoTo CheckErr
    If rst Is Nothing Then
    Else
        If rst.State = adStateOpen Then
            rst.Close
        End If
        Set rst = Nothing
    End If
    Exit Sub
CheckErr:
    LogError "Error", "CloseRS", Err.Description
End Sub

Public Function OpenTable(pstrTable As String) As ADODB.Recordset
Dim rst As New ADODB.Recordset
On Error GoTo CheckErr
    rst.Open pstrTable, ACN, adOpenDynamic, adLockPessimistic, adCmdTable
    Set OpenTable = rst
    Exit Function
CheckErr:
    LogError "Error", "OpenTable: " & pstrTable, Err.Description
End Function

Public Sub QuerySQL(ByVal pstrSQL As String, Optional ByRef plngRecordsAffected As Long)
On Error GoTo CheckErr
    ACN.BeginTrans
    ACN.Execute pstrSQL, plngRecordsAffected
    ACN.CommitTrans
    Exit Sub
CheckErr:
    ACN.RollbackTrans
    LogError "Error", "QuerySQL: " & pstrSQL, Err.Description
End Sub

'Public Function CompactDB(pstrOriginalFileName As String, pstrDestinationFileName As String) As Boolean
'    Dim oJetEngine As New JRO.JetEngine
'    Dim strSource As String
'    Dim strDestination As String
'On Error GoTo CheckErr
'    strSource = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'                "Data Source=" & pstrOriginalFileName & ";" & _
'                "Jet OLEDB:Database Password=" & gstrDatabasePassword & ";" & _
'                "Jet OLEDB:Engine Type=5;"
'
'    strDestination = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
'              "Data Source=" & pstrDestinationFileName & ";" & _
'              "Jet OLEDB:Database Password=" & gstrDatabasePassword & ";" & _
'              "Jet OLEDB:Engine Type=5;"
'    'Set ACN = Nothing
'    oJetEngine.CompactDatabase strSource, strDestination
'    Set oJetEngine = Nothing
'    CompactDB = True
'    Exit Function
'CheckErr:
'    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "CompactDB"
'    LogError "Error", "CompactDB", Err.Description
'    Set oJetEngine = Nothing
'    CompactDB = False
'End Function
