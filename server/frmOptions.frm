VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmOptions 
   Caption         =   "Server Options"
   ClientHeight    =   6330
   ClientLeft      =   6330
   ClientTop       =   3000
   ClientWidth     =   7980
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   7980
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1680
      Top             =   5700
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   5700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOptions.frx":350C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   7875
      _ExtentX        =   13891
      _ExtentY        =   9869
      _Version        =   393216
      Tabs            =   6
      Tab             =   5
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmOptions.frx":3DE6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblServerName"
      Tab(0).Control(1)=   "lblAdminName"
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(3)=   "txtServerName"
      Tab(0).Control(4)=   "txtAdminName"
      Tab(0).Control(5)=   "txtMaxClients"
      Tab(0).Control(6)=   "Check1"
      Tab(0).Control(7)=   "Check2"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "IP Administration"
      TabPicture(1)   =   "frmOptions.frx":3E02
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "User Manager"
      TabPicture(2)   =   "frmOptions.frx":3E1E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "News"
      TabPicture(3)   =   "frmOptions.frx":3E3A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2"
      Tab(3).Control(1)=   "txtNews"
      Tab(3).Control(2)=   "cmdSaveNews"
      Tab(3).Control(3)=   "cmdFonts"
      Tab(3).Control(4)=   "Check3"
      Tab(3).Control(5)=   "Check4"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "Chat Options"
      TabPicture(4)   =   "frmOptions.frx":3E56
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Message Board Manager"
      TabPicture(5)   =   "frmOptions.frx":3E72
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "Label3"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label4"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Frame4"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "txtDays"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).ControlCount=   4
      Begin VB.TextBox txtDays 
         Height          =   285
         Left            =   6420
         TabIndex        =   31
         Top             =   2100
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Group List"
         Height          =   3795
         Left            =   180
         TabIndex        =   28
         Top             =   1560
         Width           =   5235
         Begin MSComctlLib.ListView ListView4 
            Height          =   4275
            Left            =   180
            TabIndex        =   29
            Top             =   240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   7541
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Group"
               Object.Width           =   3528
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar4 
            Height          =   1140
            Left            =   3240
            TabIndex        =   30
            Top             =   240
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   2011
            ButtonWidth     =   2355
            ButtonHeight    =   1005
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "ADD"
                  Object.ToolTipText     =   "Add IP Address"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "REMOVE"
                  Object.ToolTipText     =   "Remove IP Address"
                  ImageIndex      =   3
               EndProperty
            EndProperty
         End
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Italic"
         Height          =   315
         Left            =   -72060
         TabIndex        =   26
         Top             =   3840
         Width           =   675
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Bold"
         Height          =   255
         Left            =   -72900
         TabIndex        =   25
         Top             =   3900
         Width           =   795
      End
      Begin VB.CommandButton cmdFonts 
         Caption         =   "Change Color"
         Height          =   315
         Left            =   -74880
         TabIndex        =   24
         Top             =   3900
         Width           =   1695
      End
      Begin VB.CommandButton cmdSaveNews 
         Caption         =   "Save News"
         Height          =   315
         Left            =   -70080
         TabIndex        =   23
         Top             =   4740
         Width           =   2355
      End
      Begin VB.Frame Frame3 
         Caption         =   "Channels"
         Height          =   4635
         Left            =   -74760
         TabIndex        =   20
         Top             =   780
         Width           =   5655
         Begin MSComctlLib.ListView ListView3 
            Height          =   4275
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   7541
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
               Text            =   "Channel"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Min. Privilege"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   570
            Left            =   3780
            TabIndex        =   22
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1005
            ButtonWidth     =   2355
            ButtonHeight    =   1005
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "ADD"
                  Object.ToolTipText     =   "Add IP Address"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "REMOVE"
                  Object.ToolTipText     =   "Remove IP Address"
                  ImageIndex      =   3
               EndProperty
            EndProperty
         End
      End
      Begin RichTextLib.RichTextBox txtNews 
         Height          =   2235
         Left            =   -74880
         TabIndex        =   19
         Top             =   1500
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   3942
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmOptions.frx":3E8E
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Timestamp Log Output"
         Height          =   375
         Left            =   -74760
         TabIndex        =   18
         Top             =   3540
         Width           =   2235
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Start Server on iBBS Startup"
         Height          =   315
         Left            =   -74760
         TabIndex        =   17
         Top             =   3000
         Width           =   2715
      End
      Begin VB.TextBox txtMaxClients 
         Height          =   345
         Left            =   -73320
         TabIndex        =   13
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Users"
         Height          =   4635
         Left            =   -73620
         TabIndex        =   9
         Top             =   780
         Width           =   5355
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   570
            Left            =   3180
            TabIndex        =   15
            Top             =   360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1005
            ButtonWidth     =   2805
            ButtonHeight    =   1005
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "ADD"
                  Object.ToolTipText     =   "Add User"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "REMOVE"
                  Object.ToolTipText     =   "Remove User"
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "PRIVILEGES"
                  Object.ToolTipText     =   "Change User Privalages"
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Privaleges"
            Height          =   495
            Left            =   4020
            TabIndex        =   11
            Top             =   5040
            Width           =   1335
         End
         Begin MSComctlLib.ListView ListView2 
            Height          =   4155
            Left            =   120
            TabIndex        =   10
            Top             =   300
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   7329
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            SmallIcons      =   "ImageList1"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "User Name"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Privilege"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Banned IP Address'"
         Height          =   4695
         Left            =   -74880
         TabIndex        =   7
         Top             =   720
         Width           =   4455
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   570
            Left            =   2820
            TabIndex        =   16
            Top             =   300
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1005
            ButtonWidth     =   2355
            ButtonHeight    =   1005
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "ADD"
                  Object.ToolTipText     =   "Add IP Address"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "REMOVE"
                  Object.ToolTipText     =   "Remove IP Address"
                  ImageIndex      =   3
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   4335
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   7646
            View            =   3
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "IP ADDRESS"
               Object.Width           =   5292
            EndProperty
         End
      End
      Begin VB.TextBox txtAdminName 
         Height          =   315
         Left            =   -73320
         TabIndex        =   6
         Top             =   1620
         Width           =   2955
      End
      Begin VB.TextBox txtServerName 
         Height          =   315
         Left            =   -73320
         TabIndex        =   5
         Top             =   1080
         Width           =   2955
      End
      Begin VB.Label Label4 
         Caption         =   "Number of days to hold messages:"
         Height          =   555
         Left            =   5580
         TabIndex        =   32
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   $"frmOptions.frx":3F3C
         Height          =   735
         Left            =   180
         TabIndex        =   27
         Top             =   780
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Add/Change Server News"
         Height          =   255
         Left            =   -74760
         TabIndex        =   14
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Max Clients"
         Height          =   375
         Left            =   -74760
         TabIndex        =   12
         Top             =   2220
         Width           =   1335
      End
      Begin VB.Label lblAdminName 
         Caption         =   "Administration"
         Height          =   375
         Left            =   -74760
         TabIndex        =   4
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label lblServerName 
         Caption         =   "Server Name"
         Height          =   315
         Left            =   -74760
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   6600
      TabIndex        =   1
      Top             =   5760
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ok"
      Height          =   435
      Left            =   5160
      TabIndex        =   0
      Top             =   5760
      Width           =   1275
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim itm5 As ListItem
Dim db As Database
Dim rs As Recordset
Dim rtfcolor
Dim freefiles As Long

Private Sub Command1_Click()

Load frmAddUser
frmAddUser.Show

End Sub

Private Sub cmdFonts_Click()

CommonDialog1.ShowColor
rtfcolor = CommonDialog1.Color

End Sub

Private Sub cmdSaveNews_Click()
freefiles = FreeFile

Open App.Path & "\system\news.rtf" For Output As #freefiles
Print #1, txtNews.TextRTF
Close #1

End Sub

Private Sub Command3_Click()

'Save ini changes
WritePrivateProfileString "General", "servername", txtServerName.Text, App.Path & "\settings.ini"
WritePrivateProfileString "General", "adminname", txtAdminName.Text, App.Path & "\settings.ini"
WritePrivateProfileString "General", "maxclients", txtMaxClients.Text, App.Path & "\settings.ini"
WritePrivateProfileString "General", "messagedays", txtDays.Text, App.Path & "\settings.ini"
strServerName = txtServerName.Text
strAdminName = txtAdminName.Text
strMaxClients = txtMaxClients.Text
strDays = txtDays.Text

If Check1.Value = 1 Then
    blnServerStartState = True
Else
    blnServerStartState = False
End If

If Check2.Value = 1 Then
    blnTStamp = True
Else
    blnTStamp = False
End If

'write timestamp and server states to ini
WritePrivateProfileString "General", "timestamp", blnServerStartState, App.Path & "\settings.ini"
WritePrivateProfileString "General", "serverstate", blnServerStartState, App.Path & "\settings.ini"

'write news to news file
cmdSaveNews_Click

Unload frmOptions

End Sub

Private Sub Command4_Click()

Unload frmOptions

End Sub

Private Sub Form_Load()
Dim freefiles As Long
Dim strNews As String


SSTab1.Tab = 0
'Populate general tab objects
txtServerName.Text = strServerName
txtAdminName.Text = strAdminName
txtMaxClients.Text = strMaxClients
txtDays.Text = strDays

If blnServerStartState = True Then
    Check1.Value = 1
Else
    Check1.Value = 0
End If

If blnTStamp = True Then
    Check2.Value = 1
Else
    Check2.Value = 0
End If


'Populate User list
Set db = OpenDatabase(App.Path & "\ibbs.mdb")
Set rs = db.OpenRecordset("Select * from users")

Do Until rs.EOF
    Set itm = ListView2.ListItems.Add(, , rs!Handle)
    itm.SubItems(1) = rs!Privileges
    Select Case rs!Privileges
        Case "Basic"
            itm.SmallIcon = 4
        Case "Standard"
            itm.SmallIcon = 5
        Case "Admin"
            itm.SmallIcon = 6
    End Select
    rs.MoveNext
    Set itm = Nothing
Loop

'Populate banned IP's
Set rs = db.OpenRecordset("Select * from bannedips")
Do Until rs.EOF
    Set itm = ListView1.ListItems.Add(, , rs!IPAddress)
    rs.MoveNext
    Set itm = Nothing
Loop

'Populate chat channels
Set rs = db.OpenRecordset("Select * from chatchannels")
Do Until rs.EOF
    Set itm = ListView3.ListItems.Add(, , rs!channel)
    itm.SubItems(1) = rs("min-privilage")
    rs.MoveNext
    Set itm = Nothing
Loop

'Remove db stuff from memory
rs.Close
db.Close
Set rs = Nothing
Set db = Nothing
Set itm = Nothing

'Populate news
freefiles = FreeFile

strNews = ""
Open App.Path & "\system\news.rtf" For Input As #freefiles
Do While Not EOF(freefiles)
    Line Input #freefiles, strNewsText
    strNews = strNews & strNewsText
Loop
Close #freefiles
txtNews.TextRTF = strNews

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button
    Case "ADD"
        Command1_Click
    Case "REMOVE"
        RemoveUser
End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button
    Case "ADD"
        AddIP
    Case "REMOVE"
        RemoveIP
End Select

End Sub

Private Sub RemoveIP()

'Remove the selected IP from the banned Ip address'
'Set database objects
Set db = OpenDatabase(App.Path & "\ibbs.mdb")
'Error trap for clicking remove when there are
'no items in list or nothing selected.
Set rs = db.OpenRecordset("Select * from bannedips where ipaddress ='" & ListView1.SelectedItem & "'")

If Not rs.EOF Then
    'Remove the banned IP if found
    rs.Delete
End If

rs.Close
db.Close
Set rs = Nothing
Set db = Nothing

'Remove from list
ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
Exit Sub

removeiperrors:

Select Case Err.Number

End Select

End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button
    Case "ADD"
        'Show the add new channel form
        Load frmAddChannel
        frmAddChannel.Show
    Case "REMOVE"
        'Remove selected channel
        RemoveChannel
End Select

End Sub

Private Sub RemoveChannel()

'On Error GoTo removechannelerrors
'Remove selected Channel from the database
Set db = OpenDatabase(App.Path & "\ibbs.mdb")
'Error trap for clicking remove when there are
'no items in list or nothing selected.
Set rs = db.OpenRecordset("Select * from chatchannels where channel ='" & ListView3.SelectedItem & "'")

If Not rs.EOF Then
    'Remove the chat channel if found
    rs.Delete
End If

'Close DataBase stuff
rs.Close
db.Close
Set rs = Nothing
Set db = Nothing

'Remove from list
ListView3.ListItems.Remove (ListView3.SelectedItem.Index)
Exit Sub

removechannelerrors:

Select Case Err.Number
    Case 91
        MsgBox ("You must select a Channel!")
        Exit Sub
End Select

End Sub

Private Sub RemoveUser()

'Remove selected user from the database
Set db = OpenDatabase(App.Path & "\ibbs.mdb")
'Error trap for clicking remove when there are
'no items in list or nothing selected.
Set rs = db.OpenRecordset("Select * from users where handle ='" & ListView2.SelectedItem & "'")

If Not rs.EOF Then
    'Remove the user if found
    rs.Delete
End If

'Close DataBase stuff
rs.Close
db.Close
Set rs = Nothing
Set db = Nothing

'Remove from list
ListView2.ListItems.Remove (ListView2.SelectedItem.Index)

Exit Sub

removeusererrors:

Select Case Err.Number

End Select
End Sub

Private Sub AddIP()

strIP = InputBox("Enter IP Address To Ban:", "Ban IP")

If strIP = "" Then
    'Do nothing
Else:
    'add ip
    Set db = OpenDatabase(App.Path & "\ibbs.mdb")
    Set rs = db.OpenRecordset("Select * from bannedips where ipaddress ='" & Trim(strIP) & "'")
    
    If rs.RecordCount <> 0 Then
        'Dont add the ip. That ip already in
        'db
        MsgBox ("IP Already banned!")
        Set rs = Nothing
        Set db = Nothing
    Else
        'Add the ip
        rs.AddNew
        rs!IPAddress = Trim(strIP)
        rs.Update
        
        rs.Close
        db.Close
        
        Set rs = Nothing
        Set db = Nothing
        
        Set itm = ListView1.ListItems.Add(, , strIP)
    End If
End If


End Sub

Private Sub txtNews_Click()

txtNews.SelColor = rtfcolor
If Check3.Value = 1 Then
    txtNews.SelBold = 2
Else:
    txtNews.SelBold = 0
End If

If Check4.Value = 1 Then
    txtNews.SelItalic = 2
Else:
    txtNews.SelItalic = 0
End If

End Sub
