Attribute VB_Name = "modSQLQuery"
' Version : 0.1
' Author: Poon Yip Hoon
' Created on: 26 Aug 2017
' Modified on:
' Descriptions: Commonly use database query for CRUD
' Dependencies: modDatabase

Option Explicit
'Public gstrSQL As String

Public Sub SQL_SET_Text(strField As String, strText As String, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & strField & " = '" & strText & "'"
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQL_SET_Double(strField As String, dblNumber As Double, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & strField & " = " & dblNumber
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQL_SET_Long(strField As String, lngNumber As Long, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & strField & " = " & lngNumber
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQL_SET_Boolean(strField As String, blnValue As Boolean, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & strField
    If blnValue Then
        gstrSQL = gstrSQL & " = TRUE"
    Else
        gstrSQL = gstrSQL & " = FALSE"
    End If
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQL_SET_DateTime(strField As String, strDateTime As String, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & strField & " = #" & strDateTime & "#"
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLText(strText As String, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    If blnBeginSpace = True Then
        gstrSQL = gstrSQL & " "
    End If
    gstrSQL = gstrSQL & strText
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQL_WHERE_Text(strField As String, strText As String)
    gstrSQL = gstrSQL & " WHERE " & strField & " = '" & strText & "'"
End Sub

Public Sub SQL_WHERE_Long(strField As String, lngNumber As Long)
    gstrSQL = gstrSQL & " WHERE " & strField & " = " & lngNumber
End Sub

Public Sub SQL_WHERE_Integer(strField As String, intNumber As Integer)
    gstrSQL = gstrSQL & " WHERE " & strField & " = " & intNumber
End Sub

Public Sub SQL_WHERE_Boolean(strField As String, blnBoolean As Boolean)
    gstrSQL = gstrSQL & " WHERE " & strField & " = " & blnBoolean
End Sub

Public Sub SQL_WHERE_BETWEEN(strField As String, strLeftValue As String, strRightValue As String)
    gstrSQL = gstrSQL & " WHERE " & strField & " BETWEEN " & strLeftValue & " AND " & strRightValue
End Sub

Public Sub SQL_WHERE_LIKE_Text(strField As String, strText As String)
    gstrSQL = gstrSQL & " WHERE " & strField & " LIKE '%" & strText & "%'"
End Sub

Public Sub SQL_OR_LIKE_Text(strField As String, strText As String)
    gstrSQL = gstrSQL & " OR " & strField & " LIKE '%" & strText & "%'"
End Sub

Public Sub SQL_AND_Text(strField As String, strText As String)
    gstrSQL = gstrSQL & " AND " & strField & " = '" & strText & "'"
End Sub

Public Sub SQL_ORDER_BY(strField As String, Optional blnAscending As Boolean = True)
    gstrSQL = gstrSQL & " ORDER BY " & strField
    If blnAscending = False Then
        gstrSQL = gstrSQL & " DESC"
    End If
End Sub

Public Sub SQL_INNER_JOIN(strTable1 As String, strTable2 As String, strCommonField1 As String, strCommonField2 As String)
    'SQLText "FROM " & strTable1, False
    SQLText "INNER JOIN " & strTable2, False
    SQLText "ON " & strTable1 & "." & strCommonField1 & " = " & strTable2 & "." & strCommonField2, False
End Sub

Public Sub SQL_LEFT_JOIN(strTable1 As String, strTable2 As String, strCommonField1 As String, strCommonField2 As String)
    'SQLText "FROM " & strTable1, False
    SQLText "LEFT JOIN " & strTable2, False
    SQLText "ON " & strTable1 & "." & strCommonField1 & " = " & strTable2 & "." & strCommonField2, False
End Sub

Public Sub SQL_SELECT()
    gstrSQL = "SELECT"
End Sub

Public Sub SQL_SELECT_ALL(strTable As String)
    gstrSQL = "SELECT * FROM " & strTable
End Sub

Public Sub SQL_SELECT_TOP(strField As String, strTable As String, Optional intTop As Integer = 1)
    gstrSQL = "SELECT TOP " & intTop & " " & strField & " FROM " & strTable
End Sub

Public Sub SQL_FROM(strTable As String)
    gstrSQL = gstrSQL & " FROM " & strTable
End Sub

Public Sub SQL_INSERT(strTable As String)
    gstrSQL = "INSERT INTO " & strTable & " ("
End Sub

Public Sub SQL_VALUES()
    gstrSQL = gstrSQL & ") VALUES ("
End Sub

Public Sub SQL_UPDATE(strTable As String)
    gstrSQL = "UPDATE " & strTable & " SET"
End Sub

Public Sub SQL_DELETE(strTable As String)
    gstrSQL = "DELETE FROM " & strTable
End Sub

Public Sub SQL_Comma()
    gstrSQL = gstrSQL & ","
End Sub

Public Sub SQL_Close_Bracket()
    gstrSQL = gstrSQL & ")"
End Sub

Public Sub SQLData_Text(strText As String, Optional blnEndComma As Boolean = True, Optional blnBeginSpace As Boolean = True)
    If blnBeginSpace = True Then
        gstrSQL = gstrSQL & " "
    End If
    gstrSQL = gstrSQL & "'" & CheckString(strText) & "'"
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLData_Double(dblNumber As Double, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & dblNumber
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLData_Long(lngNumber As Long, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & lngNumber
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLData_Integer(intNumber As Integer, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " " & intNumber
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLData_Boolean(blnValue As Boolean, Optional blnEndComma As Boolean = True)
    If blnValue Then
        gstrSQL = gstrSQL & " TRUE"
    Else
        gstrSQL = gstrSQL & " FALSE"
    End If
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub

Public Sub SQLData_DateTime(strDateTime As String, Optional blnEndComma As Boolean = True)
    gstrSQL = gstrSQL & " #" & strDateTime & "#"
    If blnEndComma = True Then
        gstrSQL = gstrSQL & ","
    End If
End Sub
