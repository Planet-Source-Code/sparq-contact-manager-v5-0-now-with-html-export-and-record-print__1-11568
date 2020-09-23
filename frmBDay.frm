VERSION 5.00
Begin VB.Form frmBDay 
   Caption         =   "Reminders"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   8460
   Begin VB.CommandButton cmdPic 
      Caption         =   "Show Picture"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6060
      TabIndex        =   2
      Top             =   780
      Width           =   1215
   End
   Begin VB.CommandButton cmdRec 
      Caption         =   "Show Record"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6060
      TabIndex        =   1
      Top             =   180
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2955
      ItemData        =   "frmBDay.frx":0000
      Left            =   60
      List            =   "frmBDay.frx":0002
      TabIndex        =   0
      Top             =   173
      Width           =   5955
   End
End
Attribute VB_Name = "frmBDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Table As Recordset
Dim Recs(0 To 50) As String
Dim Pics(0 To 50) As String

Private Sub cmdPic_Click()

    Dim Another As New frmPic
    Filename = Pics(List1.ListIndex)
    If Left(Filename, 1) = "~" Then
        Filename = App.Path & "\" & Right(Filename, Len(Filename) - 1)
    End If
    Another.Picture1.Picture = LoadPicture(Filename)
    Another.Width = Another.Picture1.Width + 60
    Another.Height = Another.Picture1.Height + 360
    Another.Caption = Recs(List1.ListIndex) & "'s Picture"
    Load Another
    Another.Show

End Sub

Private Sub cmdRec_Click()
    OpenContact Recs(List1.ListIndex)
End Sub

Private Sub Form_Load()
  Width = 8580
  Height = 3705
  Dim TempDate As Date
  Dim Count As Integer
  Count = 0
  Set Table = frmMain.DB.OpenRecordset("SELECT * FROM CONTACTS ORDER BY LNAME DESC")
    With Table
        .MoveFirst
        Do While Not .EOF
            TempDate = !BDayM & "/" & !BDayD & "/" & Year(Date)
            If Abs(GetDays(TempDate)) < 15 Then
                List1.AddItem "Birthday - " & !Fname & " " & !LName & "  (" & GetDays(TempDate) & " days)"
                Recs(Count) = !LName & ", " & !Fname
                Pics(Count) = !pic
                Count = Count + 1
            End If
            .MoveNext
        Loop
    End With
End Sub

Private Sub List1_Click()
 Dim Index As Integer
 Index = List1.ListIndex
    If Pics(Index) <> "" Then
        Dim Filename As String
            
            Filename = Pics(Index)
            Filename = Trim(Filename)
            If Left(Filename, 1) = "~" Then
                Filename = App.Path & Right$(Filename, Len(Filename) - 1)
            End If
            If Dir(Filename) <> "" Then
                cmdPic.Enabled = True
            Else
                cmdPic.Enabled = False
            End If
    Else
        cmdPic.Enabled = False
    End If
    
    cmdRec.Enabled = True
End Sub
