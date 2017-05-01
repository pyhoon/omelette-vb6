VERSION 5.00
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Omelette"
   ClientHeight    =   7995
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13845
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewProject 
         Caption         =   "&New Project"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenProject 
         Caption         =   "&Open Project..."
         Shortcut        =   ^O
      End
      Begin VB.Menu Separator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewProjectExplorer 
         Caption         =   "&Project Explorer"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuProject 
      Caption         =   "&Project"
   End
   Begin VB.Menu mnuFormat 
      Caption         =   "F&ormat"
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuRun 
      Caption         =   "&Run"
      Begin VB.Menu mnuRunStart 
         Caption         =   "&Start"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "Q&uery"
   End
   Begin VB.Menu mnuDiagram 
      Caption         =   "D&iagram"
   End
   Begin VB.Menu mnuAddIns 
      Caption         =   "&Add-Ins"
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Omelette..."
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    Me.WindowState = vbMaximized
End Sub

Private Sub mnuAbout_Click()
    MsgBox App.ProductName & vbCrLf & _
    "Version " & App.Major & "." & App.Minor & " Build " & App.Revision & vbCrLf & vbCrLf & _
    "Copyright" & Chr(169) & " 2017 Company Name. All rights reserved.", vbInformation, App.Title
End Sub

Private Sub mnuExit_Click()
    ' Show close dialog
    End
End Sub

Private Sub mnuFileNewProject_Click()
    With frmProjectNew
        .Show vbModal
    End With
End Sub

Private Sub mnuFileOpenProject_Click()
    With frmProjectOpen
        .Show vbModal
    End With
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show
End Sub

Private Sub mnuRunStart_Click()
    With frmLiveApp
        .Show
        .ZOrder 0
    End With
End Sub

Private Sub mnuViewProjectExplorer_Click()
    With frmWindowProject
        .Width = 4000
        .Height = 4000
        .Left = Me.Width - .Width - 300
        .Top = 0
        .Show
    End With
End Sub
