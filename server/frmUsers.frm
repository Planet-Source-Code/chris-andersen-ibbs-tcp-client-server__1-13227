VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsers 
   Caption         =   "Users Online"
   ClientHeight    =   7755
   ClientLeft      =   7050
   ClientTop       =   2040
   ClientWidth     =   5535
   Icon            =   "frmUsers.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7755
   ScaleWidth      =   5535
   Begin VB.CommandButton Command1 
      Caption         =   "Kick Selected User"
      Height          =   495
      Left            =   3180
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7395
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   13044
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ip Address"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "should also add the option to Admin message the selected user or maybe checked users. That is for later though. "
      Height          =   2535
      Left            =   3180
      TabIndex        =   2
      Top             =   1080
      Width           =   2355
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

For Each User In colUsers
    If User.Handle = ListView1.SelectedItem Then
        'Send disconect string to client. Client
        'will handle disconnect.
        a = MsgBox("Kicking " & User.Handle, vbOKOnly, "Kicking")
        frmMain.sckServer(User.SockIndex).SendData "disconnect1||"
    End If
    
Next
End Sub

Private Sub Form_Load()
Dim itm2 As ListItem

For Each User In colUsers
    DoEvents
    'Add handle and ip address of online users
    'to the list
    Set itm2 = ListView1.ListItems.Add(, , User.Handle)
    itm2.SubItems(1) = User.IPAddress
Next

Set itm2 = Nothing

End Sub
