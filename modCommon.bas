Attribute VB_Name = "modCommon"
' Version : 2.1
' Author: Poon Yip Hoon
' Modified On : 10/09/2014
' Descriptions : Commonly use functions
Option Explicit
Public gstrSQL As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Function CheckString(strInput As String) As String
    'strInput = Replace(strInput, "'", "")
    strInput = Replace(strInput, "'", "''")
    strInput = Replace(strInput, """", "")
    ' Error in SQL if use following in LIKE search
    strInput = Replace(strInput, "[", "")
    strInput = Replace(strInput, "]", "")
    strInput = Replace(strInput, "*", "")
    strInput = Replace(strInput, "&", "")
    CheckString = strInput
End Function

' Prevent Hacker's SQL Injection
' Try to key in 1=1'or'1=1 as password to hack
Public Function CheckInput(strInput As String) As String
    'Take out parts of the username that are not permitted
    strInput = Replace(strInput, "password", "", 1, -1, 1)
    strInput = Replace(strInput, "salt", "", 1, -1, 1)
    strInput = Replace(strInput, "author", "", 1, -1, 1)
    strInput = Replace(strInput, "code", "", 1, -1, 1)
    strInput = Replace(strInput, "username", "", 1, -1, 1)
    strInput = Replace(strInput, "select", "", 1, -1, 1)
    strInput = Replace(strInput, "from", "", 1, -1, 1)
    'Replace harmful SQL quotation marks with doubles
    'CheckInput = strInput 'Test
    strInput = Replace(strInput, """", "", 1, -1, 1) 'Use this
    strInput = Replace(strInput, "'", "''", 1, -1, 1)   'Use this
    'strInput = Replace(strInput, "''", " ", 1, -1, 1) 'Do not use this
    CheckInput = strInput
End Function

Public Function ConvLng(pstrInput As String) As Long
    If Trim(pstrInput) = "" Then
        ConvLng = 0
    Else
        If IsNumeric(pstrInput) = True Then
            ConvLng = CLng(pstrInput)
        Else
            ConvLng = 0
        End If
    End If
End Function

Public Function ConvDbl(pstrInput As String) As Double
    If Trim(pstrInput) = "" Then
        ConvDbl = 0
    ElseIf IsNull(pstrInput) Then
        ConvDbl = 0
    Else
        If IsNumeric(pstrInput) = True Then
            ConvDbl = CDbl(pstrInput)
        Else
            ConvDbl = 0
        End If
    End If
End Function

Public Function ConvCur(pstrInput As String) As Currency
    If Trim(pstrInput) = "" Then
        ConvCur = 0
    ElseIf IsNull(pstrInput) Then
        ConvCur = 0
    Else
        If IsNumeric(pstrInput) = True Then
            ConvCur = CCur(pstrInput)
        Else
            ConvCur = 0
        End If
    End If
End Function

Public Function FormatCurrency(pstrInput As Variant, Optional intDecimal As Integer = 2) As String
    FormatCurrency = FormatNumber(pstrInput, intDecimal, vbTrue, vbFalse, vbTrue)
End Function

Public Function FormatDate(dtDate As Date) As String
    FormatDate = Format(dtDate, "dd MMM yyyy")
End Function

Public Function FormatTime(dtDate As Date) As String
    FormatTime = Format(dtDate, "hh:mm ampm")
End Function

Public Function FormatTimeSeconds(dtDate As Date, Optional strAMPM As String = " AMPM") As String
    FormatTimeSeconds = Format(dtDate, "hh:mm:ss" & strAMPM)
End Function

Public Function FormatDateAndTime(dtDate As Date) As String
    FormatDateAndTime = Format(dtDate, "dd MMM yyyy hh:mm:ss ampm")
End Function

Public Function FormatMonthYear(dtDate As Date) As String
    FormatMonthYear = Format(dtDate, "MMM yyyy")
End Function

Public Function FormatYear(dtDate As Date) As String
    FormatYear = Format(dtDate, "yyyy")
End Function

' Get date of first day of the week for Weekly Report query
Public Function WeekDay1(dtDate As Date) As String
    Dim intDay As Integer
    Dim datDate As Date
    intDay = Weekday(dtDate, vbSunday)
    datDate = DateAdd("D", -intDay + 1, dtDate)
    WeekDay1 = FormatDate(datDate)
End Function

' Get date of last day of the week for Weekly Report query
Public Function WeekDay7(dtDate As Date) As String
    Dim intDay As Integer
    Dim datDate As Date
    intDay = Weekday(dtDate, vbSunday)
    datDate = DateAdd("D", 7 - intDay, dtDate)
    WeekDay7 = FormatDate(datDate)
End Function

' Get date of first day of the month for Monthly Report query
Public Function MonthDay1(dtDate As Date) As String
    Dim strTemp As String
    Dim datDate As Date
    strTemp = "1 "
    strTemp = strTemp & MonthName(Month(dtDate), True) & " "
    strTemp = strTemp & Year(dtDate)
    datDate = DateAdd("D", 0, strTemp)
    MonthDay1 = FormatDate(datDate)
End Function

' Get date of last day of the month for Monthly Report query
Public Function MonthDay30(dtDate As Date) As String
    Dim strTemp As String
    Dim datDate As Date
    Dim intNextMonth As Integer
    Dim intNextMonthYear As Integer
    If Month(dtDate) = 12 Then
        intNextMonth = 1
        intNextMonthYear = Year(dtDate) + 1
    Else
        intNextMonth = Month(dtDate) + 1
        intNextMonthYear = Year(dtDate)
    End If
    strTemp = "1 " & MonthName(intNextMonth, True) & " " & intNextMonthYear
    datDate = DateAdd("D", -1, strTemp)
    MonthDay30 = FormatDate(datDate)
End Function

' Get date of first day of the year for Annual Report query
Public Function YearDay1(dtDate As Date) As String
    Dim strTemp As String
    Dim datDate As Date
    strTemp = "1 "
    strTemp = strTemp & MonthName(1, True) & " "
    strTemp = strTemp & Year(dtDate)
    datDate = DateAdd("D", 0, strTemp)
    YearDay1 = FormatDate(datDate)
End Function

' Get date of last day of the year for Annual Report query
Public Function YearDay365(dtDate As Date) As String
    Dim strTemp As String
    Dim datDate As Date
    strTemp = "31 "
    strTemp = strTemp & MonthName(12, True) & " "
    strTemp = strTemp & Year(dtDate)
    datDate = DateAdd("D", 0, strTemp)
    YearDay365 = FormatDate(datDate)
End Function

Public Function FormatDigit(lngID As Long, intDigit As Integer) As String
    Dim i As Integer
    Dim strTemp As String
    Dim intLen As Integer
    intLen = Len(CStr(lngID))
    For i = 1 To (intDigit - intLen)
        strTemp = strTemp & "0"
    Next
    FormatDigit = strTemp & lngID 'Format$(lngTransID, "0000000")
End Function

Public Function FormatFixedID(strID As String, intLength As Integer) As String
    Dim i As Integer
    FormatFixedID = strID
    For i = 1 To (intLength - Len(strID))
        FormatFixedID = FormatFixedID & " "
    Next
End Function

Public Function FormatItem(strItemID As String, strItemName As String, Optional blnEnable As Boolean = True) As String
    Dim str As String
    str = "(X)" ' "Ø"
    If blnEnable = True Then
        FormatItem = strItemID & vbTab & strItemName
    Else
        FormatItem = strItemID & str & vbTab & strItemName
    End If
End Function

Public Function GetItemID(strItem As String) As String
    Dim strPart() As String
    strPart = Split(strItem, vbTab, 2)
    strItem = Replace(strPart(0), "*", "")
    GetItemID = strItem
End Function

Public Sub SetCombo(cboOutput As ComboBox, strText As String, Optional intItemData As Integer = 0)
    Dim i As Integer
    If intItemData <> 0 Then
        For i = 0 To cboOutput.ListCount - 1
            If cboOutput.ItemData(i) = intItemData Then
                cboOutput.ListIndex = i
                Exit For
            End If
        Next
    Else
        If strText <> "" Then
            For i = 0 To cboOutput.ListCount - 1
                If cboOutput.List(i) = strText Then
                    cboOutput.ListIndex = i
                    Exit For
                End If
            Next
        Else
            cboOutput.ListIndex = 0
        End If
    End If
End Sub

Public Sub SetRecord(txtOutput As TextBox, adoField As ADODB.Field, Optional blnTrim As Boolean = True)
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

Public Sub SetCheck(chkOutput As CheckBox, adoField As ADODB.Field)
    If adoField = True Then
        chkOutput.Value = vbChecked
    Else
        chkOutput.Value = vbUnchecked
    End If
End Sub

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

Public Sub WriteText(pstrFileName As String, Optional pstrNote As String)
    Open pstrFileName For Append As #1
    If pstrNote = "" Then
        Write #1, FormatDateAndTime(Now), Error
    Else
        Write #1, FormatDateAndTime(Now), Error & " - " & pstrNote
    End If
    Close #1
End Sub

Public Sub ReadText(ByVal pstrFileName As String, ByVal pintLineNo As Integer, ByRef pstrOutput As String)
    Dim i As Integer
On Error GoTo Catch
    Open pstrFileName For Input As #2
    If pintLineNo < 0 Then
            Do Until EOF(2) = True
                Input #2, pstrOutput
            Loop
        ElseIf pintLineNo > 0 Then
            For i = 0 To pintLineNo
                If Not EOF(2) Then
                    Input #2, pstrOutput
                Else
                    Exit For
                End If
            Next
        Else
            Input #2, pstrOutput
        End If
    Close
    Exit Sub
Catch:

End Sub

Public Function FileExists(strPath As String) As Boolean
    Dim lngRetVal As Long
On Error Resume Next
    lngRetVal = Len(Dir$(strPath))
    If Err Or lngRetVal = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If
End Function

Public Sub ResetControls(frm As Form)
    Dim ctl As Control
    For Each ctl In frm.Controls
        If TypeOf ctl Is TextBox Then ctl.Text = vbNullString
        If TypeOf ctl Is ComboBox Then ctl.ListIndex = -1
    Next
End Sub
