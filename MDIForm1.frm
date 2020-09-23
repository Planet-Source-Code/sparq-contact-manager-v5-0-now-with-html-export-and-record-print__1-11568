VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Sparq's Phone Book"
   ClientHeight    =   4650
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4395
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   15002
            Text            =   "E-Mail jason@alphamedia.net with Question, Comments, etc..  PLEASE VOTE."
            TextSave        =   "E-Mail jason@alphamedia.net with Question, Comments, etc..  PLEASE VOTE."
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      Begin VB.Menu mnuShowContactList 
         Caption         =   "Show / Hide Contact &List"
      End
      Begin VB.Menu mnuToDo 
         Caption         =   "Show / Hide &To Do List"
      End
   End
   Begin VB.Menu mnuExport 
      Caption         =   "&Export"
      Begin VB.Menu mnuHTMl 
         Caption         =   "As &HTML"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public DB As Database
Public ContactTable As Recordset

Private Sub MDIForm_Load()
    Set DB = OpenDatabase(App.Path & "\pbook.mdb")
    mnuShowContactList_Click
    mnuToDo_Click
    Load frmBDay
    If frmBDay.List1.ListCount < 1 Then Unload frmBDay: GoTo 10
    frmBDay.Show
10
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuHTMl_Click()
    With frmHTML
        .Left = (frmMain.Width / 2) - (.Width / 2)
        .Top = (frmMain.Height / 2) - (.Height / 2)
        .Caption = "Export as HTML"
        .txtPath = App.Path
        .Show 1
    End With
End Sub

Public Sub mnuShowContactList_Click()
    If Not frmContList.Visible Then
        Load frmContList
        frmContList.Top = 60
        frmContList.Left = 60
        frmContList.Show
    Else
        Unload frmContList
    End If
End Sub

Private Sub mnuToDo_Click()
    If Not frmToDo.Visible Then
        Load frmToDo
        frmToDo.Top = 60
        frmToDo.Height = 4725
        frmToDo.Width = 4095
        frmToDo.Left = Screen.Width - (frmToDo.Width + 160)
        frmToDo.Show
    Else
        Unload frmToDo
    End If
End Sub


