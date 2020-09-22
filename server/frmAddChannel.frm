VERSION 5.00
Begin VB.Form frmAddChannel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Chat Channel"
   ClientHeight    =   1545
   ClientLeft      =   6255
   ClientTop       =   3900
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbPrivilage 
      Height          =   315
      Left            =   1260
      TabIndex        =   5
      Text            =   "Basic"
      Top             =   540
      Width           =   2355
   End
   Begin VB.TextBox txtChannel 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   3
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "Channel Name"
      Height          =   270
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "Privilage Req."
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmAddChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub cmdCancel_Click()

Unload frmAddChannel

End Sub

Private Sub cmdOK_Click()

If txtChannel.Text = "" Then
    MsgBox ("You must enter a channel name!")
    Exit Sub
End If

'add user
Set db = OpenDatabase(App.Path & "\ibbs.mdb")
Set rs = db.OpenRecordset("Select * from chatchannels where channel ='" & Trim(txtChannel) & "'")


If rs.RecordCount <> 0 Then
    'Dont add the channel. That name already in
    'db
    MsgBox ("Chat Channel already exist's!")
    Set rs = Nothing
    Set db = Nothing
Else
    'Add the channel
    rs.AddNew
    rs!Channel = Trim(txtChannel)
    rs("min-privilage") = Trim(cbPrivilage.Text)
    rs.Update
    
    rs.Close
    db.Close
    
    Set rs = Nothing
    Set db = Nothing
    
    Set itm = frmOptions.ListView3.ListItems.Add(, , txtChannel.Text)
    itm.SubItems(1) = cbPrivilage.Text
    
    Unload frmAddChannel
End If

End Sub


Private Sub Form_Load()

cbPrivilage.AddItem "Basic"
cbPrivilage.AddItem "Standard"
cbPrivilage.AddItem "Admin"


End Sub
