VERSION 5.00
Begin VB.Form frmAddUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add User"
   ClientHeight    =   1965
   ClientLeft      =   8025
   ClientTop       =   5055
   ClientWidth     =   3750
   Icon            =   "frmAddUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1160.987
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbPrivilage 
      Height          =   315
      Left            =   1275
      TabIndex        =   6
      Text            =   "Basic"
      Top             =   960
      Width           =   2355
   End
   Begin VB.TextBox txtUserName 
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
      TabIndex        =   4
      Top             =   1380
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1320
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "Privilage"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset

Private Sub cmdCancel_Click()

Unload frmAddUser

End Sub

Private Sub cmdOK_Click()

If txtUserName.Text = "" Then
    MsgBox ("You must enter a user name!")
    Exit Sub
End If

'add user
Set db = OpenDatabase(App.Path & "\ibbs.mdb")
Set rs = db.OpenRecordset("Select * from users where handle ='" & Trim(txtUserName) & "'")

If rs.RecordCount <> 0 Then
    'Dont add the user. That name already in
    'db
    a = MsgBox("User by this name already exist's!")
    Set rs = Nothing
    Set db = Nothing
Else
    'Add the user
    rs.AddNew
    rs!Handle = Trim(txtUserName)
    rs!Password = Trim(txtPassword)
    rs!Privileges = Trim(cbPrivilage.Text)
    rs.Update
    
    rs.Close
    db.Close
    
    Set rs = Nothing
    Set db = Nothing
    
    Set itm = frmOptions.ListView2.ListItems.Add(, , txtUserName)
    itm.SubItems(1) = cbPrivilage.Text
    
    Select Case cbPrivilage.Text
        Case "Basic"
            itm.SmallIcon = 4
        Case "Standard"
            itm.SmallIcon = 5
        Case "Admin"
            itm.SmallIcon = 6
    End Select
        
    Unload frmAddUser
End If

End Sub

Private Sub Form_Load()

cbPrivilage.AddItem "Basic"
cbPrivilage.AddItem "Standard"
cbPrivilage.AddItem "Admin"

End Sub
