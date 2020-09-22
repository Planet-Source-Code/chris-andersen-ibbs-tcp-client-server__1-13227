VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "iBBS Client"
   ClientHeight    =   7785
   ClientLeft      =   1635
   ClientTop       =   1950
   ClientWidth     =   12315
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "News"
            Object.ToolTipText     =   "Latest Server News"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Chat"
            Object.ToolTipText     =   "Join Chat Now"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Files"
            Object.ToolTipText     =   "Download Files"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MessageBoard"
            Object.ToolTipText     =   "Read and Write Messages"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mailbox"
            Object.ToolTipText     =   "Check your personal mail"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Browser"
            Object.ToolTipText     =   "Browse the Web"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7410
      Width           =   12315
      _ExtentX        =   21722
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "Server Name:"
            TextSave        =   "Server Name:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "Admin:"
            TextSave        =   "Admin:"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   5460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":350C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   10380
      Top             =   300
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   10920
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnudisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu mnuconnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuview 
      Caption         =   "View"
      Begin VB.Menu mnunews 
         Caption         =   "Latest News"
      End
      Begin VB.Menu mnuchat 
         Caption         =   "Chat"
      End
      Begin VB.Menu mnufiles 
         Caption         =   "Files"
      End
      Begin VB.Menu mnuim 
         Caption         =   "Instant Message"
      End
      Begin VB.Menu mnumb 
         Caption         =   "Message Forum"
      End
      Begin VB.Menu mnuBrowse 
         Caption         =   "Web Browser"
      End
      Begin VB.Menu mnumail 
         Caption         =   "Check Mailbox"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About iBBS"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the client is fairly straight forward.
Dim strUserList As String
Dim IMWindowFound As Boolean
Dim itm1 As ListItem
Dim lngIcon As Long


Private Sub MDIForm_Load()

With sckClient
    .RemoteHost = strHost
    .RemotePort = "1001"
    .Connect
End With

'Set encryption key. Change this for your own version(s).
key = "ãÎ•úËžâÀ¾€=Â1n´"

Load Userlist
Userlist.Show

'Preset this variable so that incoming chat data doesnt print to the chat window
'until it has been opened
strChatFormState = "Closed"

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

'Close iBBS client
End

End Sub

Private Sub mnuBrowse_Click()

Load frmBrowser
frmBrowser.Show

End Sub

Private Sub mnuchat_Click()

'Load and size chat form
Load chatform
With chatform
    .Height = 8520
    .Width = 6705
    .Show
End With

End Sub

Private Sub mnuconnect_Click()

MDIForm1.Hide
Load frmLogin
frmLogin.Show

End Sub

Private Sub mnudisconnect_Click()

sckClient.Close

End Sub

Private Sub mnufiles_Click()

'Load and size file form
Load fileform
With fileform
    .Height = 8010
    .Width = 4170
    .Show
End With

End Sub

Private Sub mnumail_Click()

Timer1.Interval = 0
'load and size mailbox form
Load frmMail
With frmMail
    .Height = 5550
    .Width = 6705
    .Show
End With

Timer1.Interval = 10000
End Sub

Private Sub mnumb_Click()

'load and size Message board form
Load frmMBoard
With frmMBoard
    .Height = 6645
    .Width = 6390
    .Show
End With


End Sub

Private Sub mnunews_Click()

Load frmNews

Timer1.Interval = 0
'Send news request
sckClient.SendData "news1||"
'Size news form
With frmNews
    .Height = 3060
    .Width = 9255
    .Show
End With
Timer1.Interval = 10000

End Sub

Private Sub mnuquit_Click()

End

End Sub

Private Sub sckClient_DataArrival(ByVal bytesTotal As Long)

Dim strSendCode As String
Dim vntArray As Variant
Dim strText As String
Dim nItems As Integer
Dim n As Integer

Timer1.Interval = 0
sckClient.GetData strSendCode, vbString

' split function will be used to parse items contained in a string,
' and delimitted by ||
' The Split function returns a variant array containing each parsed item
' as an element in the array. VB6 and up only

' use split function to parse it
vntArray = Split(strSendCode, "||")

' how many items were parsed?
nItems = UBound(vntArray)
  
'Compare vntarray(0)(the "send code") to determine appropriate action(s)
Select Case vntArray(0)
    Case "channellist1"
        'populate channel list if chat form is open
        If strChatFormState <> "Closed" Then
            For Y = 1 To nItems
                chatform.cbChannel.AddItem vntArray(Y)
            Next Y
        End If
    Case "disconnect1"
        'Disconnect Message from server. Close the connecetion
        'For if user was kicked.
        sckClient.Close
        a = MsgBox("You have been kicked by the Server Admin!", vbCritical, "Connection Lost")
    Case "admin1"
        'Incoming Admin message
        strtest = MsgBox(vntArray(1), vbCritical, "Message from the Administrator")
        
    Case "news2"
        'incomng news
        frmNews.Text1.TextRTF = vntArray(1)
        
    Case "connect1"
        'Determine if the server is allowing access
        Select Case vntArray(1)
            Case "logonyes"
                strHandle = strUser
                MDIForm1.Show
                MDIForm1.WindowState = vbMaximized
                Unload frmLogin
                StatusBar1.Panels(1).Text = "CONNECTED"
                StatusBar1.Panels(2).Text = "Server Name: " & vntArray(3)
                StatusBar1.Panels(3).Text = "Admin: " & vntArray(4)
            Case "logonno"
                a = MsgBox("Login Incorrect!", vbCritical, "Login Failed")
            Case "logonbanned"
                a = MsgBox("You have been banned from this server! If you have been banned in error, contact the Server's Admin.", vbCritical, "Logon Failed")
            Case "ok"
                sckClient.SendData "ibbslogin1||" & strUser & "||" & strPass
        End Select
        
    Case "imcode2"
        'Incoming Instant message
        Dim strIMMessage As String '
        Dim strIMer As String 'The person the IM is from
        
        strIMMessage = vntArray(1)
        'RC4 decode message
        strIMMessage = RC4(strIMMessage, key)
        
        strIMer = vntArray(2)
        
        'First check if window for that IM'ing User is already open
        'If it is, send the text to the appropriate IM Window
        'If not create a new IM window then send the text to it
        IMWindowFound = False
        
        
        'checking the Forms collection for the approriate IM window
        For Each frm In Forms
            'DoEvents
            If frm.Caption = vntArray(2) Then
                'IM window already opened from this user.
                'Give that window the focus
                frm.SetFocus
                'Output the text to it
                With frm.Text1
                    .SelColor = vbRed
                    .SelBold = 2
                    .SelText = "<" & strIMer & ">"
                    .SelColor = vbBlack
                    .SelText = strIMMessage & vbCrLf
                End With
                
                IMWindowFound = True
                Exit For
            End If
        Next
        
        If IMWindowFound = False Then
            'New User IM. Create a new IM Window
            Load IMForm(IMNumber)
            
            With IMForm(IMNumber)
                With IMForm(IMNumber).Text1
                    .SelColor = vbRed
                    .SelBold = 2
                    .SelText = "<" & strIMer & ">"
                    .SelColor = vbBlack
                    .SelText = strIMMessage & vbCrLf
                End With
                .Caption = strIMer
                .Height = 4350
                .Width = 6090
                .Show
            End With
            
            IMNumber = IMNumber + 1
        End If
        
        IMWindowFound = False
        
    Case "imcode3"
        For Each frm In Forms
            'DoEvents
            If frm.Caption = vntArray(1) Then
                'IM window already opened from this user.
                'Give that window the focus
                frm.SetFocus
                'Output the user not found text
                With frm.Text1
                    .SelColor = vbBlack
                    .SelBold = 2
                    .SelText = "User no longer online!!!"
                End With
                
                Exit For
            End If
        Next
        
    Case "chatcode2"
        Dim strChatHandle As String
        Dim strMessage As String
        
        'Handle incoming chat info
        strChatHandle = vntArray(1)
        strMessage = vntArray(2)
        
        With chatform.Text1
            If strChatHandle = strHandle Then
                .SelColor = vbGreen
            Else
                .SelColor = vbRed
            End If
            
            .SelBold = 2
            
            'Checking for emote
            If InStr(1, strMessage, "/me") = 0 Then
                .SelText = "<" & strChatHandle & ">"
            Else
                .SelColor = vbBlue
                strMessage = Replace(strMessage, "/me", "")
                .SelText = strChatHandle
            End If
            
            .SelBold = 1
            .SelColor = vbBlack
            .SelText = strMessage & vbCrLf
        End With
        
 
    Case "filelistcode2"
        nItems = UBound(vntArray)

        ' display each file available for download on the server
       
        fileform.ListView1.ListItems.Clear
        For n = 1 To nItems - 1
            vntarray2 = Split(vntArray(n), "\/")
            
            
            Set itm = fileform.ListView1.ListItems.Add(, , vntarray2(1))
            itm.SubItems(1) = vntarray2(0)
            'Change icon for file in list depending on its extension
            Select Case LCase(Right(vntarray2(1), 3))
                Case "txt"
                    itm.SmallIcon = 1
                Case "mp3"
                    itm.SmallIcon = 2
                Case "wav", "mid"
                    itm.SmallIcon = 3
                Case "zip"
                    itm.SmallIcon = 4
                Case "jpg", "gif", "bmp"
                    itm.SmallIcon = 5
                Case "vbs"  'Show a danger icon since this is a danger file to download.
                    itm.SmallIcon = 6
                Case "doc"
                    itm.SmallIcon = 8
                Case Else
                    itm.SmallIcon = 7
            End Select
            
        Next n
        Set itm = Nothing
        
    Case "mbmessagescode2"
        'Implementation on hold for now pending better way to do it
        nItems = UBound(vntArray)

         'display each parsed item
      
        For n = 1 To nItems - 1
            vntarray2 = Split(vntArray(n), "\/")

            Set itm = frmMBoard.ListView2.ListItems.Add(, , vntarray2(0))
            itm.SubItems(1) = vntarray2(1)
            itm.SubItems(2) = vntarray2(2)
        Next n
        
        Set itm = Nothing
        
    Case "mbgroupscode2"
        'display group list
        nItems = UBound(vntArray)
                
        For n = 1 To nItems - 1
            vntarray2 = Split(vntArray(n), "\/")

            Set itm = frmMBoard.ListView1.ListItems.Add(, , vntarray2(0))
            itm.SubItems(2) = vntarray2(1)
        Next n
        
        Set itm = Nothing
    
    Case "mbmessageexpired2"
        'remove expired message
        MsgBox ("This message has expired!")
        
        frmMBoard.ListView2.ListItems.Remove frmMBoard.ListView2.SelectedItem.Index
        
    Case "userlistcode2"
        'display user lists
        nItems = UBound(vntArray)
        
        If strChatFormState <> "Closed" Then
            chatform.List1.Clear
        End If
  
        Userlist.userlist1.ListItems.Clear

        For n = 1 To nItems - 1
            vntarray2 = Split(vntArray(n), "@@")
            strUserList = strUserList & vntarray2(1) & "||"
            'add users in same chat channel to chat user list
            If strChatFormState <> "Closed" And vntarray2(1) = strChannel Then
               chatform.List1.AddItem vntarray2(0)
            End If
            'add all users to main userlist
            Set itm1 = Userlist.userlist1.ListItems.Add(, , vntarray2(0))
            itm1.SubItems(1) = vntarray2(1)
        Next n
        Set itm1 = Nothing
            
    Case "mailcode2"
        'parse mails subject, from, and id into mailbox listview
        nItems = UBound(vntArray)

        For n = 1 To nItems - 1
            vntarray2 = Split(vntArray(n), "@@")
            
            Set itm1 = frmMail.ListView1.ListItems.Add(, , vntarray2(0))
            itm1.SubItems(1) = vntarray2(1)
            itm1.SubItems(2) = vntarray2(2)
        Next n
        
        Set itm1 = Nothing
        
    Case "mailmessagecode2"
        
        frmMail.txtMail.SelText = vntArray(1)
    
    Case "mbgetmessagecode2"
        
        frmMBoard.RichTextBox1.SelText = vntArray(1)
    
    Case "invalidmailcode1"
        
        MsgBox ("Last sent mail message was undeliverable because of an invalid recipient!")
        
        
End Select

Timer1.Interval = 5000

End Sub

Private Sub Timer1_Timer()

'Send Userlist request.
If sckClient.State = 7 Then
    sckClient.SendData "userlist1"
Else:
    StatusBar1.Panels(1).Text = "DISCONNECTED"
    MsgBox ("Connection to Server Lost!")
    Timer1.Interval = 0
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.key
    Case "News"
        mnunews_Click
    Case "Chat"
        mnuchat_Click
    Case "Files"
        mnufiles_Click
    Case "MessageBoard"
        mnumb_Click
    Case "Mailbox"
        mnumail_Click
    Case "Browser"
        mnuBrowse_Click
    End Select

End Sub
