VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00303030&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4245
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSplash 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Version"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5400
      TabIndex        =   1
      Top             =   2400
      Width           =   525
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "ProductName"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   870
      Left            =   1650
      TabIndex        =   0
      Top             =   1560
      Width           =   4275
   End
   Begin VB.Shape shpBottom 
      BorderColor     =   &H00505050&
      FillColor       =   &H00505050&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   -240
      Top             =   3960
      Width           =   7935
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With lblProductName
        .Caption = App.Title
        .BackStyle = vbTransparent
    End With
    
    With lblVersion
        .Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
        .BackStyle = vbTransparent
        .Left = lblProductName.Left + lblProductName.Width - lblVersion.Width
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Main
End Sub

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub tmrSplash_Timer()
    Unload Me
End Sub
