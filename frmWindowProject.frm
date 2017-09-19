VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWindowProject 
   Caption         =   "Project"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3765
   Icon            =   "frmWindowProject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   3765
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
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5318
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

Private Sub Form_Load()
    'RefreshTreeview
End Sub

Public Sub RefreshTreeview(strProjectType As String)  ' Temporary use project type
    Dim strLabel As String
    Dim strKey As String
    Dim nod As Node
On Error GoTo Catch
    With tvwFiles
        .Enabled = False
        .Nodes.Clear
        
        If gstrProjectName = "" Then Exit Sub
        
        strKey = "G1"
        strLabel = gstrProjectName & " (" & gstrProjectName & ".vbp)" ' ProjectTitle (ProjectTitle.vbp)
        Set nod = .Nodes.Add(, tvwChild, strKey, strLabel, 1)
        nod.Bold = True
        nod.Expanded = True
        
        ' Temporary show STANDARD project
        If strProjectType = "STANDARD" Then
            strKey = "Forms"
            strLabel = "Forms"
            Set nod = .Nodes.Add("G1", tvwChild, strKey, strLabel, 2)
            nod.Expanded = True
                    
            strKey = "Item1"
            strLabel = "frmForm1 (frmForm1.frm)"
            Set nod = .Nodes.Add("Forms", tvwChild, strKey, strLabel, 3)
            '.Font.Name = "Verdana" 'Verdana
            '.Font.Charset = 0
            '.Font.Size = 10
            .Enabled = True
        End If
    End With
Exit Sub
Catch:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "RefreshTreeview"
    LogError "Error", "RefreshTreeview/frmWindowProject", Err.Description
End Sub
