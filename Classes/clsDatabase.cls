VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mstrErrorDesc As String 'local copy
Public Property Let ErrorDesc(ByVal strErrorDesc As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ErrorDesc = 5
    mstrErrorDesc = strErrorDesc
End Property


Public Property Get ErrorDesc() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ErrorDesc
    ErrorDesc = mstrErrorDesc
End Property

Public Sub ProjectQuerySQL(ByVal pstrSQL As String, Optional ByRef plngRecordsAffected As Long)
On Error GoTo CatchErr
    With gconProject
        .BeginTrans
        .Execute pstrSQL, plngRecordsAffected
        .CommitTrans
    End With
    Exit Sub
CatchErr:
    ErrorDesc = Err.Description
    On Error Resume Next
    gconProject.RollbackTrans
    LogError "Error", "ProjectQuerySQL: " & pstrSQL, Err.Description
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
    ErrorDesc = Err.Description
    gconItem.RollbackTrans
    LogError "Error", "ItemQuerySQL: " & pstrSQL, Err.Description
End Sub

