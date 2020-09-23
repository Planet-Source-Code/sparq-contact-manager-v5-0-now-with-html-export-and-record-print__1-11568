VERSION 5.00
Begin VB.Form frmContList 
   BackColor       =   &H00404000&
   BorderStyle     =   0  'None
   Caption         =   "   Contacts"
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   2715
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   4320
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   2715
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2100
      TabIndex        =   4
      Top             =   4500
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   4500
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2340
      TabIndex        =   2
      Top             =   0
      Width           =   330
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
      Caption         =   " CONTACTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "frmContList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Set frmMain.DB = OpenDatabase(App.Path & "\pbook.mdb")
  Set frmMain.ContactTable = frmMain.DB.OpenRecordset("SELECT * FROM CONTACTS ORDER BY LNAME ASC")
  LoadContacts
End Sub

Sub LoadContacts()
    List1.Clear
    With frmMain.ContactTable
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            List1.AddItem !LName & ", " & !Fname
            .MoveNext
        Loop
    End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Visible = False
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then FormDrag Me
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub Label2_Click()
    frmMain.mnuShowContactList_Click
End Sub

Private Sub Label3_Click()
    With frmMain.ContactTable
        .AddNew
        !Fname = "Contact"
        !LName = "New"
        .Update
    End With
    LoadContacts
    OpenContact "New, Contact"
End Sub

Private Sub Label4_Click()
    If List1.ListIndex < 0 Then Exit Sub
    Dim Answer As Integer
        On Error GoTo Err
        Answer = MsgBox("Delete: " & List1.Text & "?", vbYesNo + vbQuestion, "Delete")
        If Answer = vbYes Then
            With frmMain.ContactTable
                .MoveFirst
                Do Until !LName & ", " & !Fname = List1.Text
                    .MoveNext
                Loop
                .Delete
            End With
        End If
        LoadContacts
        Exit Sub
Err:
    MsgBox "Error Deleting: " & List1.Text, vbCritical, "Error"
End Sub

Private Sub List1_DblClick()
    OpenContact List1.Text
End Sub
