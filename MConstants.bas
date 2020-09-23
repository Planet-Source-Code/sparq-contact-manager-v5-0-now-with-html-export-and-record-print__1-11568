Attribute VB_Name = "MConstants"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 5
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function tapiRequestMakeCall& Lib "TAPI32.DLL" (ByVal DestAddress$, ByVal AppName$, ByVal CalledParty$, ByVal Comment$)
Private Const TAPIERR_NOREQUESTRECIPIENT = -2&
Private Const TAPIERR_REQUESTQUEUEFULL = -3&
Private Const TAPIERR_INVALDESTADDRESS = -4&

Public Sub ExportFilesToHTML(Optional Path As String, Optional Title As String)
    If Path = "" Then Path = App.Path & "\"
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    If Title = "" Then Title = "My Contacts"
    Set frmMain.ContactTable = frmMain.DB.OpenRecordset("SELECT * FROM CONTACTS ORDER BY LNAME ASC")
    With frmMain.ContactTable
        .MoveFirst
        Open Path & "Index.htm" For Output As #1
            Print #1, "<HTML>"
            Print #1, "<HEAD>"
            Print #1, "<TITLE>" & Title & "</TITLE>"
            Print #1, "</HEAD>"
            Print #1, "<BODY BGColor=""#FFFFFF"" Text=""#000000"">"
            Print #1, "<H2>" & Title & "</H2>"
            Print #1, "<UL>"
            Do While Not .EOF
                Print #1, "<LI><A HREF=""" & !LName & !Fname & ".htm"">" & !LName & ", " & !Fname & "</A></LI>"
                .MoveNext
            Loop
            Print #1, "</UL>"
            Print #1, "</BODY>"
            Print #1, "</HTML>"
        Close #1
        .MoveFirst
        frmHTML.PBar1.Min = 0
        frmHTML.PBar1.Max = .RecordCount
        frmHTML.PBar1.Value = 0
        Do While Not .EOF
            Open Path & !LName & !Fname & ".htm" For Output As #1
                On Error Resume Next
                Print #1, "<HTML>"
                Print #1, "<HEAD>"
                Print #1, "<TITLE>" & !LName & " " & !Fname & "</TITLE>"
                Print #1, "</HEAD>"
                Print #1, "<BODY BGColor=""#FFFFFF"" Text=""#000000"">"
                Print #1, "<H2>" & !LName & " " & !Fname & "</H2>"
                If !Address1 <> "" Then Print #1, !Address1 & "<BR>"
                If !Address2 <> "" Then Print #1, !Address2 & "<BR>"
                If Not (!State = "" And !City = "" And !Zip = "") Then
                    Print #1, !City & ", " & !State & " " & !Zip & "<BR><BR>"
                End If
                If !Phone1 <> "" Then Print #1, "<B>Phone</B>: " & !Phone1 & "<BR>"
                If !Phone2 <> "" Then Print #1, "<B>Phone</B>: " & !Phone2 & "<BR>"
                If !Cell <> "" Then Print #1, "<B>Cell</B>: " & !Cell & "<BR>"
                If !fax <> "" Then Print #1, "<B>Fax</B>: " & !fax & "<BR>"
                Print #1, "<BR><B>Category</B>: " & frmContact.cmbCat.List(!Cat) & "<BR>"
                If !EMail <> "" Then Print #1, "<B>E-Mail</B>: " & "<A HREF=""mailto:" & !EMail & """>" & !EMail & "</A><BR>"
                If !URL <> "" Then Print #1, "<B>URL</B>: " & "<A HREF=""" & !URL & """>" & !URL & "</A><BR>"
                Print #1, "<B>Birthday</B>: " & Format(!BDayM, "00") & "/" & Format(!BDayD, "00") & "/" & Format(!BDayY, "00") & "<BR>"
                If !Notes <> "" Then
                    Print #1, "<BR><B>Notes</B>:<BR>"
                    Print #1, !Notes & "<BR>"
                End If
                Print #1, "<BR><A HREF=""Index.htm"">Back to Contacts</A>"
                Print #1, "</BODY>"
                Print #1, "</HTML>"
                frmHTML.PBar1.Value = frmHTML.PBar1.Value + 1
            Close #1
            .MoveNext
        Loop
    End With
    Dim Answer As Integer
      Answer = MsgBox("Export Done! Do you want to view now?", vbYesNo + vbQuestion, "Done.")
      If Answer = vbYes Then
          ShellExecute frmContact.hwnd, "open", Path & "Index.htm", vbNullString, vbNullString, SW_SHOW
      End If
End Sub

Public Sub Dial(Frm As Form, Num As String)
  Dim buff As String
  Dim nResult As Long
    nResult = tapiRequestMakeCall&(Trim$(Num), CStr(Frm.Caption), Frm.txtLName & ", " & Frm.txtFName, "")
    If nResult <> 0 Then
        buff = "Error dialing number : "
        Select Case nResult
               Case TAPIERR_NOREQUESTRECIPIENT
                    buff = buff & "No Windows Telephony dialing application is running and none could be started."
               Case TAPIERR_REQUESTQUEUEFULL
                    buff = buff & "The queue of pending Windows Telephony dialing requests is full."
               Case TAPIERR_INVALDESTADDRESS
                    buff = buff & "The phone number is Not valid."
               Case Else
                    buff = buff & "Unknown error."
               End Select
    End If
End Sub


Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Public Function FormatNumber(Text As String) As String
  Dim X As Integer
  Dim TempNum As String
  Dim CurLet As String
    For X = 1 To Len(Text)
        CurLet = Mid(Text, X, 1)
        If IsNumeric(CurLet) Then TempNum = TempNum & CurLet
    Next X
    FormatNumber = TempNum
End Function

Public Sub OpenContact(Name As String)
  Dim X As Integer
  Dim Another As New frmContact
  Dim YearDiff As Integer
  
    'Check if record is already open by searching the captions of all loaded forms.
    For X = 0 To Forms.Count - 1
        'If so, Exit sub
        If Forms(X).Caption = "Contacts - " & Name Then Forms(X).SetFocus: Exit Sub
    Next X
    
    
    With frmMain.ContactTable
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            If !LName & ", " & !Fname = Name Then
                Exit Do
            Else
                .MoveNext
            End If
        Loop
        
        Dim BDate As Date
        On Error Resume Next
        Another.Visible = False
        Another.Width = 6570
        Another.Height = 6780
        Another.Caption = "Contacts - " & Name
        Another.txtFName = !Fname
        Another.txtLName = !LName
        Another.txtPhone1 = !Phone1
        Another.txtPhone2 = !Phone2
        Another.txtCell = !Cell
        Another.txtFax = !fax
        Another.txtAdd1 = !Address1
        Another.txtAdd2 = !Address2
        Another.txtCity = !City
        Another.txtState = UCase(!State)
        Another.txtZip = !Zip
        Another.txtNotes = !Notes
        Another.txtNotes.TabIndex = 0
        Another.txtEmail = !EMail
        Another.txtPic = !pic
        
        Dim Filename As String
        If Left(Another.txtPic, 1) = "~" Then
            Filename = App.Path & "\" & Right(Another.txtPic, Len(Another.txtPic) - 1)
        Else
            Filename = Another.txtPic
        End If

        If Filename = "" Or Dir(Filename) = "" Then
            Another.lblPic.Visible = False
        Else
            Another.lblPic.Visible = True
        End If
        
        Another.txtURL = !URL
        Another.cmbBDayM.ListIndex = Val(!BDayM) - 1
        Another.cmbBDayD.ListIndex = Val(!BDayD) - 1
        YearDiff = Year(Date) - !BDayY
        Another.cmbBDayY.ListIndex = Another.cmbBDayY.ListCount - (YearDiff + 1)
        Another.Tag = Name
        Another.cmbCat.ListIndex = !Cat
        BDate = !BDayM & "/" & !BDayD & "/" & Year(Date)
        Another.lblDays = "Days until BDay: " & GetDays(BDate)
        Load Another
        Another.Visible = True
        Another.Changes = False
        Another.SSTab1.Tab = 0
        Another.Show
    End With
    Exit Sub
End Sub

Public Function GetDays(BDate As Date) As Integer
    If DateDiff("d", Date, BDate) < 1 Then
        BDate = Month(BDate) & "/" & Day(BDate) & "/" & Year(BDate) + 1
    End If
    GetDays = DateDiff("d", Date, BDate)
End Function


Public Sub PrintRecord(Name As String)
  Dim Found As Boolean
    Found = False
    With frmMain.ContactTable
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            If !LName & ", " & !Fname = Name Then
                Found = True
                Exit Do
            Else
                .MoveNext
                Found = False
            End If
        Loop
        
        If Not (Found) Then MsgBox "Record not found", vbExclamation, "Error": Exit Sub
        
        Printer.ScaleMode = vbInches
        Printer.CurrentX = 0
        Printer.CurrentY = 0
        Printer.FontSize = 18
        Printer.Print Name
        Printer.FontSize = 12
        Printer.CurrentY = Printer.CurrentY + 0.5
        Printer.CurrentX = 0
        
        If !Address1 <> "" Then
            Printer.Print !Address1
            Printer.CurrentY = Printer.CurrentY + 0.08
            Printer.CurrentX = 0
        End If
        If !Address2 <> "" Then
            Printer.Print !Address2
            Printer.CurrentY = Printer.CurrentY + 0.08
            Printer.CurrentX = 0
        End If
        If Not (!City = "" And !State = "" And !Zip = "") Then
            Printer.Print !City & ", " & !State & " "; !Zip
            Printer.CurrentY = Printer.CurrentY + 0.08
            Printer.CurrentX = 0
        End If
        
        If !Phone1 <> "" Then
            Printer.Print "Phone 1: " & !Phone1
            Printer.CurrentY = Printer.CurrentY + 0.08
            Printer.CurrentX = 0
        End If
        If !Phone2 <> "" Then
            Printer.Print "Phone 2: " & !Phone2
            Printer.CurrentY = Printer.CurrentY + 0.08
            Printer.CurrentX = 0
        End If
        If !Cell <> "" Then
            Printer.Print "Cell: " & !Cell
            Printer.CurrentY = Printer.CurrentY + 0.08
            Printer.CurrentX = 0
        End If
        If !fax <> "" Then
            Printer.Print "Fax: " & !fax
            Printer.CurrentY = Printer.CurrentY + 0.08
            Printer.CurrentX = 0
        End If
        If !EMail <> "" Then
            Printer.Print "E-Mail: " & !EMail
            Printer.CurrentY = Printer.CurrentY + 0.08
            Printer.CurrentX = 0
        End If
        If !URL <> "" Then
            Printer.Print "URL: " & !URL
            Printer.CurrentY = Printer.CurrentY + 0.08
            Printer.CurrentX = 0
        End If
        
        Printer.Print "Category: " & frmContact.cmbCat.List(Val(!Cat))
        Printer.CurrentY = Printer.CurrentY + 0.08
        Printer.CurrentX = 0
        
        Printer.Print "Birthday: " & Format(!BDayM, "00") & "/" & Format(!BDayD, "00") & "/" & Format(!BDayY, "00")
        Printer.CurrentY = Printer.CurrentY + 0.08
        Printer.CurrentX = 0
        
        If !Notes <> "" Then
            Printer.Print "Notes: " & !Notes
        End If
        
        Printer.EndDoc
    End With
End Sub
