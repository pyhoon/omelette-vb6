VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProjectManage 
   Caption         =   "Manage Project"
   ClientHeight    =   4320
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8730
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      MaxLength       =   255
      TabIndex        =   7
      Top             =   3600
      Width           =   5535
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      MaxLength       =   255
      TabIndex        =   6
      Top             =   3240
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      MaxLength       =   255
      TabIndex        =   5
      Top             =   2880
      Width           =   5535
   End
   Begin MSComctlLib.ListView lvwProjects 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3945
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Project File"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Project Path"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Project Name"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
   End
End
Attribute VB_Name = "frmProjectManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
