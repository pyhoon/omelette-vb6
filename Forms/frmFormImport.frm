VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportTemplate 
   Caption         =   "Import File to Template"
   ClientHeight    =   2910
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboMethod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmFormImport.frx":0000
      Left            =   1680
      List            =   "frmFormImport.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   4935
   End
   Begin VB.ComboBox cboTemplate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmFormImport.frx":0069
      Left            =   1680
      List            =   "frmFormImport.frx":0079
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   4935
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   1200
      Width           =   615
   End
   Begin MSComDlg.CommonDialog cdgSourceFile 
      Left            =   3120
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtSourceFile 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      MaxLength       =   255
      TabIndex        =   7
      Top             =   1200
      Width           =   4335
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      MaxLength       =   255
      TabIndex        =   10
      Top             =   1680
      Width           =   4935
   End
   Begin VB.CommandButton cmdImport 
      Appearance      =   0  'Flat
      Caption         =   "Import"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Source File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1200
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "File Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "File Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1200
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Template"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1200
   End
End
Attribute VB_Name = "frmImportTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdImport_Click()
Dim DB As New OmlDatabase
Dim SQ As New OmlSQLBuilder
Dim strOutput As String
Dim FF As Integer
On Error GoTo Catch
Try:
    If Trim(txtSourceFile.Text) = "" Or Trim(txtFileName.Text) = "" Then
        MsgBox "Please select a File", vbExclamation, "Import Form"
        Exit Sub
    End If
    DB.DataPath = gstrMasterDataPath & "\"
    DB.DataFile = gstrMasterDataFile
    DB.OpenMdb
    FF = FreeFile
    Open txtSourceFile.Text For Input As #FF
    Do While Not EOF(FF)
        Line Input #FF, strOutput
        strOutput = Replace(strOutput, "'", "''")
        SQ.INSERT "Code"
        SQ.SOB "CodeText"
        SQ.SQL "FileName"
        SQ.SQL "MethodName"
        SQ.SQL "TemplateName"
        SQ.SQL "CreatedDate"
        SQ.SCB "CreatedBy"
        SQ.VALUES
        SQ.VOB strOutput
        SQ.VAL Trim(txtFileName.Text)
        SQ.VAL Trim(cboMethod.ItemData(cboMethod.ListIndex))
        SQ.VAL Trim(cboTemplate.Text)
        SQ.DAT Now
        SQ.VCB "Aeric"
        'Debug.Print SQ.Text
        DB.Execute SQ.Text
        If DB.ErrorDesc <> "" Then
            MsgBox "Error: " & DB.ErrorDesc, vbExclamation, "Import Form"
            Close
            DB.CloseMdb
            Exit Sub
        End If
    Loop
    Close
    DB.CloseMdb
    MsgBox "Form has beem imported!", vbInformation, "Import Form"
    Exit Sub
Catch:
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "Import Form"
    LogError "Error", "ImportForm(" & Trim(txtFileName.Text), Err.Description
    Close
    DB.CloseMdb
End Sub

Private Sub Form_Load()
    cdgSourceFile.Filter = "Project Files (*.vbp)|*.vbp|Form Files (*.frm)|*.frm|Module Files (*.bas)|*.bas|Class Files (*.cls)|*.cls"
    cboTemplate.ListIndex = 0
    cboMethod.ListIndex = 0
End Sub

Private Sub cmdBrowse_Click()
Try:
    On Error GoTo Catch
    cdgSourceFile.CancelError = True
    cdgSourceFile.InitDir = App.Path
    If txtSourceFile.Text <> "" Then
        If FileExists(txtSourceFile.Text) Then
            cdgSourceFile.InitDir = Dir(txtSourceFile.Text)
            cdgSourceFile.FileName = txtSourceFile.Text
        End If
    End If
    ' Display the open dialog box
    cdgSourceFile.ShowOpen
    txtSourceFile.Text = cdgSourceFile.FileName
    Dim strTemp() As String
    strTemp = Split(txtSourceFile.Text, "\")
    txtFileName.Text = strTemp(UBound(strTemp))
    Exit Sub
Catch:
    If Err.Number = 32755 Then Exit Sub
    MsgBox Err.Number & " - " & Err.Description, vbExclamation, "Browse"
    LogError "Error", "cmdBrowse", Err.Description
End Sub
