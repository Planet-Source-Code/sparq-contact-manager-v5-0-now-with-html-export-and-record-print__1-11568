VERSION 5.00
Begin VB.Form frmToDo 
   BackColor       =   &H00404000&
   BorderStyle     =   0  'None
   Caption         =   "To Do List"
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstToDo 
      Appearance      =   0  'Flat
      Height          =   4305
      ItemData        =   "frmToDo.frx":0000
      Left            =   0
      List            =   "frmToDo.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   180
      Width           =   4095
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
      Left            =   1657
      TabIndex        =   3
      Top             =   4500
      Width           =   780
   End
   Begin VB.Shape Border 
      Height          =   495
      Left            =   1440
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3720
      TabIndex        =   1
      Top             =   0
      Width           =   330
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000040&
      Caption         =   " TO DO"
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
      TabIndex        =   2
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmToDo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Table As Recordset
Dim Active As Boolean

Private Sub Form_Load()
    LoadItems
End Sub

Private Sub LoadItems()
    Active = False
    Dim Count As Integer
    Count = 0
    lstToDo.Clear
    Set Table = frmMain.DB.OpenRecordset("SELECT * FROM ToDo ORDER BY ITEM DESC")
    With Table
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            lstToDo.AddItem !item, Count
            lstToDo.Selected(Count) = !Done
            .MoveNext
            Count = Count + 1
        Loop
    End With
    Active = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Form_Resize()
    With lstToDo
        .Top = Label1.Height
        .Left = 0
        .Width = Width
        .Height = Height - (Label1.Height * 2)
    End With
    With Border
        .Top = 0
        .Left = 0
        .Width = Width
        .Height = Height
    End With
    Label3.Left = (Width / 2) - (Label3.Width / 2)
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub Label3_Click()
    Dim Answer As String
    Answer = InputBox("Enter To Do Item:", "New Item")
    If Answer = "" Then Exit Sub
    Set Table = frmMain.DB.OpenRecordset("SELECT * FROM ToDo ORDER BY DATE DESC")
    With Table
        .AddNew
        !item = Answer
        !Done = False
        !Date = Date
        .Update
    End With
    LoadItems
End Sub

Private Sub lstToDo_ItemCheck(item As Integer)
    If Active = False Then Exit Sub
    Dim Checked As Integer
    If lstToDo.Selected(item) Then Checked = True
    If Checked Then
        CheckItem item
    Else
        UnCheckItem item
    End If
End Sub


Sub CheckItem(item As Integer)
  
  
    Set Table = frmMain.DB.OpenRecordset("SELECT * FROM ToDo ORDER BY ITEM DESC")
    With Table
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            If !item = lstToDo.List(item) Then
                .Edit
                !Done = True
                .Update
                Exit Do
            Else
                .MoveNext
            End If
        Loop
        Dim Asnwer As Integer
          Answer = MsgBox("Delete From List?" & vbCrLf & lstToDo.List(item), vbYesNo + vbQuestion, "Item Done")
          If Answer = vbNo Then Exit Sub
          .Delete
          LoadItems
    End With
End Sub

Sub UnCheckItem(item As Integer)
    With Table
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            If !item = lstToDo.List(item) Then
                .Edit
                !Done = False
                .Update
                Exit Do
            Else
                .MoveNext
            End If
        Loop
        LoadItems
    End With
End Sub
