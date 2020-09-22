VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMBoard 
   Caption         =   "Message Board"
   ClientHeight    =   6240
   ClientLeft      =   6435
   ClientTop       =   1800
   ClientWidth     =   6270
   Icon            =   "frmMBoard.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   6270
   Begin VB.CommandButton cmdNewMsg 
      Caption         =   "New Message"
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   120
      Width           =   1515
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2655
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4683
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMBoard.frx":08CA
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   2100
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2566
      View            =   3
      Sorted          =   -1  'True
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
         Text            =   "Message #"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1515
      Left            =   0
      TabIndex        =   0
      Top             =   540
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2672
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Message Group"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Number of Messages"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Group #"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmMBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNewMsg_Click()

Load frmMBmessage
With frmMBmessage
    .Height = 4710
    .Width = 6705
    .Show
End With

End Sub

Private Sub Form_Load()

'Send code to get groups
MDIForm1.sckClient.SendData "mbgroupscode1||"
'MDIForm1.sckClient.SendData "mbmessagescode1||"

End Sub


Private Sub Form_Unload(Cancel As Integer)

Set frmMBoard = Nothing

End Sub

Private Sub ListView1_DblClick()

ListView2.ListItems.Clear

'Get message list for selected group
MDIForm1.sckClient.SendData "mbmessagescode1||" & ListView1.SelectedItem.SubItems(2)

'Save selected group id for new message post
lngMBGroup = ListView1.SelectedItem.SubItems(2)
strMBGroup = ListView1.SelectedItem

End Sub

Private Sub ListView2_DblClick()

RichTextBox1.Text = ""
'Get selected message
MDIForm1.sckClient.SendData "mbgetmessagecode1||" & ListView2.SelectedItem.SubItems(2)

'save selected message id for new reply
lngMessage = ListView2.SelectedItem.SubItems(2)

End Sub
