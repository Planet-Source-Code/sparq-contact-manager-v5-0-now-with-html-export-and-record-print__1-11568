VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHTML 
   Caption         =   "Form1"
   ClientHeight    =   2130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PBar1 
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   1860
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox txtTitle 
      Height          =   255
      Left            =   180
      TabIndex        =   4
      Text            =   "My Contacts"
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Export!"
      Default         =   -1  'True
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   1320
      Width           =   1155
   End
   Begin VB.TextBox txtPath 
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   300
      Width           =   4335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Main Title:"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Export Path:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   870
   End
End
Attribute VB_Name = "frmHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    ExportFilesToHTML txtPath, txtTitle
    Unload Me
End Sub

