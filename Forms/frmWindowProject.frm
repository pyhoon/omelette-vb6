VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWindowProject 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Project"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   3765
   Icon            =   "frmWindowProject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList istIcons 
      Left            =   2880
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWindowProject.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWindowProject.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmWindowProject.frx":0B40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwFiles 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5530
      _Version        =   393217
      Indentation     =   573
      Style           =   7
      ImageList       =   "istIcons"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmWindowProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim strProjectName As String

Private Sub Form_Load()
    'RefreshTreeview
    ReadItemsData
End Sub

Public Sub ReadItemsData()
    Dim strLabel As String
    Dim strKey As String
    Dim nod As Node
    Dim rst As adodb.Recordset
    Dim i As Integer
On Error GoTo Catch
'    If gstrProjectName = "" Then
'        ItemOpenDB
'        Set rst = ItemOpenTable("Project")
'        rst.MoveFirst
'        gstrProjectName = rst!ProjectName
'        CloseRS rst
'        ItemCloseDB
'    End If
    With tvwFiles
        .Enabled = False
        .Nodes.Clear
        If gstrProjectName = "" Then Exit Sub
                
        'If strProjectFolder = "" Then strProjectFolder = App.Path & "\Projects\"
        'gstrProjectDataPath = strProjectFolder & strProjectName & "\" & gstrProjectData
        ItemOpenDB
        If gconItem Is Nothing Then ' gconItem.State <> adStateOpen Then
            'ItemCloseDB
            Exit Sub
        End If
        Set rst = ItemOpenTable("Project")
        If rst Is Nothing Then
            'CloseRS rst
            Exit Sub
        End If
        With rst
            If .EOF Then
                CloseRS rst
                Exit Sub
            End If
            .MoveFirst
            strKey = "G1"
            strLabel = rst!ProjectName & " (" & !ProjectFile & ")" ' ProjectTitle (ProjectTitle.vbp)
            Set nod = tvwFiles.Nodes.Add(, tvwChild, strKey, strLabel, 1)
            nod.Bold = True
            nod.Expanded = True
            CloseRS rst
        End With
        ItemCloseDB
                       
        strKey = "Forms"
        strLabel = "Forms"
        Set nod = .Nodes.Add("G1", tvwChild, strKey, strLabel, 2)
        nod.Expanded = True
                    
        ItemOpenDB
        Set rst = ItemOpenTable("Forms")
        With rst
            '.MoveFirst
            While Not .EOF
                i = i + 1
                strKey = "Item" & i
                strLabel = !FormName & " (" & !FormFile & ")"
                Set nod = tvwFiles.Nodes.Add("Forms", tvwChild, strKey, strLabel, 3)
                .MoveNext
            Wend
            CloseRS rst
        End With
        ItemCloseDB
            
        'strKey = "Item1"
        'strLabel = "frmForm1 (frmForm1.frm)"
        'Set nod = .Nodes.Add("Forms", tvwChild, strKey, strLabel, 3)
        '.Font.Name = "Verdana" 'Verdana
        '.Font.Charset = 0
        '.Font.Size = 10
        .Enabled = True
    End With
Exit Sub
Catch:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "ReadItemsData"
    LogError "Error", "ReadItemsData/frmWindowProject", Err.Description
End Sub

'Public Sub RefreshTreeview(Optional strProjectType As String = "STANDARD")  ' Temporary use project type
'    Dim strLabel As String
'    Dim strKey As String
'    Dim nod As Node
'On Error GoTo Catch
'    With tvwFiles
'        .Enabled = False
'        .Nodes.Clear
'
'        If gstrProjectName = "" Then Exit Sub
'
'        strKey = "G1"
'        strLabel = gstrProjectName & " (" & gstrProjectName & ".vbp)" ' ProjectTitle (ProjectTitle.vbp)
'        Set nod = .Nodes.Add(, tvwChild, strKey, strLabel, 1)
'        nod.Bold = True
'        nod.Expanded = True
'
'        ' Temporary show STANDARD project
'        If strProjectType = "STANDARD" Then
'            strKey = "Forms"
'            strLabel = "Forms"
'            Set nod = .Nodes.Add("G1", tvwChild, strKey, strLabel, 2)
'            nod.Expanded = True
'
'            strKey = "Item1"
'            strLabel = "frmForm1 (frmForm1.frm)"
'            Set nod = .Nodes.Add("Forms", tvwChild, strKey, strLabel, 3)
'            '.Font.Name = "Verdana" 'Verdana
'            '.Font.Charset = 0
'            '.Font.Size = 10
'            .Enabled = True
'        End If
'    End With
'Exit Sub
'Catch:
'    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "RefreshTreeview"
'    LogError "Error", "RefreshTreeview/frmWindowProject", Err.Description
'End Sub
