VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContact 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contacts"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   6450
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   14737632
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmContact.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Line4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDays"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label6"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label7"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label10"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label12"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label13"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label11"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line3"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lblPic"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtNotes"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtFName"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtLName"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmbCat"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtAdd1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtAdd2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtCity"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtState"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtZip"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtPhone1"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtPhone2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtFax"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtCell"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtEmail"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtURL"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmbBDayM"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "cmbBDayD"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "cmbBDayY"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).ControlCount=   38
      TabCaption(1)   =   "Picture"
      TabPicture(1)   =   "frmContact.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CD1"
      Tab(1).Control(1)=   "Picture1"
      Tab(1).Control(2)=   "Command3"
      Tab(1).Control(3)=   "txtPic"
      Tab(1).Control(4)=   "Label15"
      Tab(1).ControlCount=   5
      Begin MSComDlg.CommonDialog CD1 
         Left            =   -72540
         Top             =   2340
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   495
         Left            =   -74820
         ScaleHeight     =   435
         ScaleWidth      =   1155
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   315
         Left            =   -69420
         TabIndex        =   37
         Top             =   480
         Width           =   435
      End
      Begin VB.TextBox txtPic 
         Height          =   285
         Left            =   -74820
         TabIndex        =   36
         Top             =   480
         Width           =   5355
      End
      Begin VB.ComboBox cmbBDayY 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2880
         Width           =   1095
      End
      Begin VB.ComboBox cmbBDayD 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2880
         Width           =   675
      End
      Begin VB.ComboBox cmbBDayM 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox txtURL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   3540
         Width           =   6135
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   2940
         Width           =   3375
      End
      Begin VB.TextBox txtCell 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   5040
         TabIndex        =   15
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtFax 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   3720
         TabIndex        =   14
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox txtPhone2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   5040
         TabIndex        =   13
         Top             =   1380
         Width           =   1215
      End
      Begin VB.TextBox txtPhone1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   3720
         TabIndex        =   12
         Top             =   1380
         Width           =   1215
      End
      Begin VB.TextBox txtZip 
         Height          =   315
         Left            =   2280
         TabIndex        =   11
         Top             =   2340
         Width           =   1215
      End
      Begin VB.TextBox txtState 
         Height          =   315
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   10
         Top             =   2340
         Width           =   375
      End
      Begin VB.TextBox txtCity 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   2340
         Width           =   1575
      End
      Begin VB.TextBox txtAdd2 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   1740
         Width           =   3375
      End
      Begin VB.TextBox txtAdd1 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   1380
         Width           =   3375
      End
      Begin VB.ComboBox cmbCat 
         Height          =   315
         Left            =   3720
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtLName 
         Height          =   315
         Left            =   1860
         TabIndex        =   5
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtFName 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtNotes 
         Height          =   1515
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   4200
         Width           =   6135
      End
      Begin VB.Label lblPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View Contact's Picture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   4320
         TabIndex        =   40
         Top             =   3930
         Width           =   1935
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Hint: A Tilde (~) to the begining of the path is the equivilent of App.Path"
         Height          =   195
         Left            =   -74820
         TabIndex        =   39
         Top             =   780
         Width           =   5040
      End
      Begin VB.Line Line3 
         X1              =   3600
         X2              =   3600
         Y1              =   480
         Y2              =   3480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Birthday:"
         Height          =   195
         Left            =   3720
         TabIndex        =   35
         Top             =   2580
         Width           =   615
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   3720
         X2              =   6240
         Y1              =   2415
         Y2              =   2415
      End
      Begin VB.Line Line1 
         X1              =   3720
         X2              =   6240
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Address:        (Double - Click to visit URL)"
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   3300
         Width           =   3315
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail Address:           (Double - Click to E-Mail)"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2700
         Width           =   3330
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cell / Pager:"
         Height          =   195
         Left            =   5040
         TabIndex        =   32
         Top             =   1740
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax:"
         Height          =   195
         Left            =   3720
         TabIndex        =   31
         Top             =   1740
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Phone(s):      (Double - Click to dial)"
         Height          =   195
         Left            =   3720
         TabIndex        =   30
         Top             =   1140
         Width           =   2490
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         Height          =   195
         Left            =   2280
         TabIndex        =   29
         Top             =   2100
         Width           =   225
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "State:"
         Height          =   195
         Left            =   1740
         TabIndex        =   28
         Top             =   2100
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "City:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   2100
         Width           =   300
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   1140
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category:"
         Height          =   195
         Left            =   3720
         TabIndex        =   25
         Top             =   480
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         Height          =   195
         Left            =   1860
         TabIndex        =   24
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   3960
         Width           =   465
      End
      Begin VB.Label lblDays 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   4440
         TabIndex        =   21
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   3600
         X2              =   3600
         Y1              =   480
         Y2              =   3480
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print Record"
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   5970
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update Record"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   5970
      Width           =   1635
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<--- Now Working"
      Height          =   195
      Left            =   1980
      TabIndex        =   41
      Top             =   6060
      Width           =   1245
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5
Public Changes As Boolean

Private Sub cmbBDayD_Click()
    Changes = True
    UpdateDays
End Sub

Private Sub cmbBDayM_Click()
    Changes = True
    UpdateDays
End Sub

Private Sub cmbBDayY_Click()
    Changes = True
End Sub

Private Sub cmbCat_Click()
    Changes = True
End Sub

Private Sub Command1_Click()
    If Changes = True Then UpdateMe: MsgBox "Updated!", vbExclamation, "Updated"
End Sub

Private Sub Command2_Click()
    PrintRecord (txtLName & ", " & txtFName)
    Exit Sub
    MsgBox "PrintRecord isnt working yet - I am having problems testing it" & vbCrLf & _
           "because my printer is a P.O.S.  (if you dont know what POS means," & vbCrLf & _
           "It is not good :)"
End Sub

Private Sub Command3_Click()
  Dim Filename
    CD1.CancelError = True
    CD1.Filter = "Bitmaps (*.bmp)|*.bmp|Gif Files (*.gif)|*.gif|Jpeg Files (*.jpg)|*.jpg|"
    On Error GoTo Err
    CD1.ShowOpen
    
    If LCase(Left(CD1.Filename, Len(App.Path))) = LCase(App.Path) Then
        Filename = "~" & Right(CD1.Filename, Len(CD1.Filename) - Len(App.Path))
    Else
        Filename = CD1.Filename
    End If
    txtPic = Filename
    
Err:
End Sub

Private Sub Form_Load()
    FillDates
End Sub

Sub UpdateDays()
    If cmbBDayM.ListCount < 1 Then Exit Sub
    If cmbBDayD.ListCount < 1 Then Exit Sub
    If cmbBDayM.ListIndex < 0 Then Exit Sub
    If cmbBDayD.ListIndex < 0 Then Exit Sub
    
    Dim TDate As Date
On Error GoTo Err
    TDate = Left(cmbBDayM.Text, 2) & "/" & cmbBDayD.Text & "/00"
    lblDays = "Days until BDay: " & GetDays(TDate)
Err:
End Sub

Sub FillDates()
  Dim X As Integer
  Dim TempDate As Date
    For X = 1 To 12
        TempDate = Format(X, "00") & "/01/00"
        cmbBDayM.AddItem Format(X, "00") & "- " & Format(TempDate, "mmm")
    Next X
    For X = 1 To 31
        cmbBDayD.AddItem Format(X, "00")
    Next X
    
    For X = 1930 To Year(Date)
        cmbBDayY.AddItem X
    Next X
    cmbCat.AddItem "Friend"
    cmbCat.AddItem "Family"
    cmbCat.AddItem "Co-Worker"
    cmbCat.AddItem "General"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Changes Then
      Dim Answer As Integer
        Answer = MsgBox("Do you want to update this record?", vbYesNoCancel + vbQuestion, "Update")
        If Answer = vbYes Then
            UpdateMe
            Unload Me
        ElseIf Answer = vbCancel Then
            Cancel = True
        End If
    End If
End Sub

Sub UpdateMe()
    With frmMain.ContactTable
        .MoveFirst
        Do While Not .EOF
            If !LName & ", " & !Fname = Tag Then
                Exit Do
            Else
                .MoveNext
            End If
        Loop
        On Error Resume Next
        .Edit
        If txtFName <> "" Then !Fname = txtFName
        If txtLName <> "" Then !LName = txtLName
        If txtPhone1 <> "" Then !Phone1 = txtPhone1
        If txtPhone2 <> "" Then !Phone2 = txtPhone2
        If txtCell <> "" Then !Cell = txtCell
        If txtFax <> "" Then !fax = txtFax
        If txtAdd1 <> "" Then !Address1 = txtAdd1
        If txtAdd2 <> "" Then !Address2 = txtAdd2
        If txtCity <> "" Then !City = txtCity
        If txtState <> "" Then !State = UCase(txtState)
        If txtZip <> "" Then !Zip = txtZip
        If txtNotes <> "" Then !Notes = txtNotes
        If txtEmail <> "" Then !EMail = txtEmail
        If txtURL <> "" Then !URL = txtURL
        If txtPic <> "" Then
            !pic = txtPic
        Else
            !pic = " "
        End If
        If cmbCat.ListIndex > -1 Then !Cat = cmbCat.ListIndex
        If cmbBDayM.ListIndex > -1 Then !BDayM = Left(cmbBDayM.Text, 2)
        If cmbBDayD.ListIndex > -1 Then !BDayD = cmbBDayD.ListIndex + 1
        If cmbBDayY.ListIndex > -1 Then !BDayY = cmbBDayY.Text
        .Update
    End With
    frmContList.LoadContacts
    Changes = False
End Sub

Private Sub lblPic_Click()
    OpenPic
End Sub

Private Sub txtAdd1_Change()
    Changes = True
End Sub

Private Sub txtAdd2_Change()
    Changes = True
End Sub

Private Sub txtCell_Change()
    Changes = True
End Sub

Private Sub txtCell_DblClick()
    Dial Me, FormatNumber(txtCell)
End Sub

Private Sub txtCity_Change()
    Changes = True
End Sub

Private Sub txtEmail_Change()
    Changes = True
End Sub

Private Sub txtEmail_DblClick()
  Dim AtSpot As Integer
  Dim DotSpot As Integer
    AtSpot = InStr(0, txtEmail, "@")
    AtSpot = InStr(0, txtEmail, ".")
    If AtSpot = 0 Or DotSpot = 0 Then Exit Sub
    ShellExecute hwnd, "open", "mailto:" & txtEmail, vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub txtFax_Change()
    Changes = True
End Sub

Private Sub txtFax_DblClick()
    Dial Me, FormatNumber(txtFax)
End Sub

Private Sub txtFName_Change()
    Changes = True
End Sub

Private Sub txtLName_Change()
    Changes = True
End Sub

Private Sub txtNotes_Change()
    Changes = True
End Sub

Private Sub txtPhone1_Change()
    Changes = True
End Sub

Private Sub txtPhone1_DblClick()
    Dial Me, FormatNumber(txtPhone1)
End Sub

Private Sub txtPhone2_Change()
    Changes = True
End Sub

Private Sub txtPhone2_DblClick()
    Dial Me, FormatNumber(txtPhone2)
End Sub

Private Sub txtPic_Change()
    On Error GoTo Err
    Changes = True
    If txtPic = "" Or txtPic = " " Then Picture1.Visible = False
    If InStr(1, txtPic, ".") = 0 Then GoTo Err
    
    Filename = txtPic
    If Left$(txtPic.Text, 1) = "~" Then Filename = App.Path & "\" & Right$(txtPic, Len(txtPic) - 1)
    
    If Dir(Filename) <> "" Then
        Picture1.Picture = LoadPicture(Filename)
        Picture1.Visible = True
    Else
        Picture1.Visible = False
    End If
    lblPic.Visible = Picture1.Visible
    Changes = True
Err:
End Sub

Private Sub txtState_Change()
    Changes = True
End Sub

Private Sub txtURL_Change()
    Changes = True
End Sub

Private Sub txtURL_DblClick()
    If txtURL = "" Then Exit Sub
    ShellExecute hwnd, "open", txtURL, vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub txtZip_Change()
    Changes = True
End Sub

Sub OpenPic()
    Dim Another As New frmPic
    Another.Picture1.Picture = Picture1.Picture
    Another.Width = Another.Picture1.Width + 60
    Another.Height = Another.Picture1.Height + 360
    Another.Caption = txtFName & " " & txtLName & "'s Picture"
    Load Another
    Another.Show
End Sub

