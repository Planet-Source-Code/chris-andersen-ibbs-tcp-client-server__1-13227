VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMail 
   Caption         =   "iBBS Mailbox"
   ClientHeight    =   5145
   ClientLeft      =   7470
   ClientTop       =   4425
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   6585
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh Mail"
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Top             =   60
      Width           =   1875
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete Mail Message"
      Height          =   315
      Left            =   2220
      TabIndex        =   3
      Top             =   60
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send New Mail"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   1875
   End
   Begin RichTextLib.RichTextBox txtMail 
      Height          =   3015
      Left            =   60
      TabIndex        =   1
      Top             =   2100
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5318
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMail.frx":0000
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1575
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Subject"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "From"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mail #"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Load frmWriteMail
With frmWriteMail
    .Height = 4425
    .Width = 5265
    .Show
End With

End Sub

Private Sub Command2_Click()

txtMail.Text = ""
'delete selected mail message
MDIForm1.sckClient.SendData "deletemailcode1||" & ListView1.SelectedItem.SubItems(2)
ListView1.ListItems.Remove ListView1.SelectedItem.Index

End Sub

Private Sub Command3_Click()
txtMail.Text = ""
ListView1.ListItems.Clear

'send get mail string
MDIForm1.sckClient.SendData "maillistcode1||" & strHandle

End Sub

Private Sub Form_Load()

'send get mail string
MDIForm1.sckClient.SendData "maillistcode1||" & strHandle

End Sub



Private Sub Form_Unload(Cancel As Integer)
Set frmMail = Nothing

End Sub

Private Sub ListView1_DblClick()

txtMail.Text = ""

'get mail message
MDIForm1.sckClient.SendData "mailmessagecode1||" & ListView1.SelectedItem.SubItems(2)

End Sub
