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
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
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
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    Me.WindowState = vbMaximized
End Sub

Private Sub mnuExit_Click()
    ' Show close dialog
    End
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show
End Sub

Private Sub mnuProject_Click()

End Sub

Private Sub mnuRun_Click()
    frmLiveApp.Show
    frmLiveApp.ZOrder 0
End Sub
