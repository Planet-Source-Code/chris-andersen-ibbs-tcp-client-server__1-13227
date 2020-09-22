VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "iBBS Server"
   ClientHeight    =   9270
   ClientLeft      =   4005
   ClientTop       =   1005
   ClientWidth     =   7170
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   618
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   Begin RichTextLib.RichTextBox txtClientOutput 
      Height          =   7695
      Left            =   60
      TabIndex        =   6
      Top             =   1440
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   13573
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   3420
      Left            =   4620
      TabIndex        =   5
      Top             =   1500
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   6033
      ButtonWidth     =   3863
      ButtonHeight    =   1005
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Start/Stop Server"
            Key             =   "Start/Stop"
            Object.ToolTipText     =   "Start or Stop The Server"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Server Options"
            Key             =   "Options"
            Object.ToolTipText     =   "Change Server Options"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close iBBS Server"
            Key             =   "Close"
            Object.ToolTipText     =   "Shutdown iBBS Server"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "User's Online"
            Key             =   "Users"
            Object.ToolTipText     =   "View/Kick Online Users"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Send Admin Message"
            Key             =   "AdminMessage"
            Object.ToolTipText     =   "Send a quick message to all users online"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Dump To Log"
            Object.ToolTipText     =   "Dump Log to an external RTF file"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock ftpSock1 
      Index           =   0
      Left            =   5940
      Top             =   4620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4980
      Top             =   4620
   End
   Begin MSWinsockLib.Winsock sckServer 
      Index           =   0
      Left            =   5460
      Top             =   4620
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame frStats 
      Caption         =   "Server Statistics"
      Height          =   3795
      Left            =   4620
      TabIndex        =   0
      Top             =   5220
      Width           =   2415
      Begin VB.Label Label2 
         Caption         =   "Server Runtime:"
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Max Users:"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   780
         Width           =   2115
      End
      Begin VB.Label lblusernm 
         Height          =   255
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lbluser 
         Caption         =   "Users Online:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1035
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6420
      Top             =   4500
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
            Picture         =   "frmMain.frx":04F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2858
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3132
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   600
      Picture         =   "frmMain.frx":3A0C
      Stretch         =   -1  'True
      Top             =   60
      Width           =   5835
   End
   Begin VB.Line Line3 
      X1              =   304
      X2              =   476
      Y1              =   340
      Y2              =   340
   End
   Begin VB.Line Line2 
      X1              =   300
      X2              =   300
      Y1              =   96
      Y2              =   608
   End
   Begin VB.Line Line1 
      X1              =   4
      X2              =   472
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSendCode As String
Dim lngSock As Long
Dim lngUsers As Long
Dim blnServerOn As Boolean
Dim blnServerPaused As Boolean
Dim blnSendDone As Boolean
Dim strMode As String
Dim strSubject As String
Dim strFrom As String
Dim strDate As String
Dim strReply As String
Dim strID As String
Dim db2 As Database
Dim rs2 As Recordset
Dim strMBString As String
Dim strFileStatus As String
Dim lngFtp As Long
Dim intBuffer As Integer
Dim lngBytesXfer As Long
Dim fFile As Long
Dim strAdminMessage As String
Dim sdata As String
Dim ldatalen As Long
Dim db As Database
Dim rs As Recordset
Dim strChannels As String

Private Sub cmdOptions_Click()

Load frmOptions
frmOptions.Show

End Sub

Private Sub Command1_Click()
'turn server on and off

If blnServerOn = True Then
    'if server on, turn it off
    sckServer(0).Close
    Dim X As Long
    For X = 1 To sckServer.UBound
        sckServer(X).Close
    Next X
    
    'Remove all users from collection
    For Each User In colUsers
        colUsers.Remove CStr(User.SockIndex)
    Next
    
    blnServerOn = False
    With txtClientOutput
        .SelColor = vbRed
        .SelBold = 1
        .SelText = TimeStamp & "Server Status: Off" & vbCrLf
        .SelBold = 0
    End With
    
Else
    'if server off, turn it on
    sckServer(0).Listen
    blnServerOn = True
    With txtClientOutput
        .SelColor = vbBlue
        .SelBold = 1
        .SelText = TimeStamp & "Server Status: On" & vbCrLf
    End With
End If


End Sub

Private Sub Command2_Click()
'close program
End

End Sub

Private Sub Command3_Click()

Load frmUsers
frmUsers.Show

End Sub

Private Sub Command4_Click()

strAdminMessage = InputBox("Enter Message:", "Admin Message")

If strAdminMessage = "" Then
    'Do nothing
Else:
    'send message
    With txtClientOutput
        .SelColor = vbBlack
        .SelBold = 2
        .SelText = TimeStamp & "Admin Message Sent: "
        .SelColor = vbRed
        .SelText = strAdminMessage & vbCrLf
        .SelBold = 0
    End With
    
    strAdminMessage = "admin1||" & strAdminMessage
    'Send the admin message to all clients.
    For X = 1 To sckServer.UBound
        If sckServer(X).State <> 7 Then
        Else
            sckServer(X).SendData strAdminMessage
            DoEvents
        End If
    Next X
End If

End Sub

Private Sub LogDump()
Dim frfile As Long
Dim filetimedate As String

'Dump data from Output text box to an rtf file.
frfile = FreeFile
filetimedate = Format(Now(), "(mm dd yyyy) (#hh-mm-ss)")

Open App.Path & "\logs\logdump" & filetimedate & ".rtf" For Output As #frfile
Print #frfile, txtClientOutput.TextRTF
Close #frfile

End Sub

Private Sub Form_Load()

'Get ini settings
sdata = Space$(255)
ldatalen = GetPrivateProfileString("General", "servername", "", sdata, Len(sdata), App.Path & "\settings.ini")
sdata = Left$(sdata, ldatalen)
strServerName = sdata

sdata = Space$(255)
ldatalen = GetPrivateProfileString("General", "adminname", "", sdata, Len(sdata), App.Path & "\settings.ini")
sdata = Left$(sdata, ldatalen)
strAdminName = sdata

sdata = Space$(255)
ldatalen = GetPrivateProfileString("General", "maxclients", "", sdata, Len(sdata), App.Path & "\settings.ini")
sdata = Left$(sdata, ldatalen)
strMaxClients = sdata

sdata = Space$(255)
ldatalen = GetPrivateProfileString("General", "serverstate", "", sdata, Len(sdata), App.Path & "\settings.ini")
sdata = Left$(sdata, ldatalen)
blnServerStartState = sdata

sdata = Space$(255)
ldatalen = GetPrivateProfileString("General", "timestamp", "", sdata, Len(sdata), App.Path & "\settings.ini")
sdata = Left$(sdata, ldatalen)
blnTStamp = sdata

sdata = Space$(255)
ldatalen = GetPrivateProfileString("General", "messagedays", "", sdata, Len(sdata), App.Path & "\settings.ini")
sdata = Left$(sdata, ldatalen)
strDays = sdata

Label1.Caption = "Max Users: " & strMaxClients

'Load Chat channels into string
Set db = OpenDatabase(App.Path & "\ibbs.mdb")
Set rs = db.OpenRecordset("Select * from chatchannels")

Do Until rs.EOF
    strChannels = strChannels & "||" & rs!channel
    rs.MoveNext
Loop

rs.Close
db.Close
Set rs = Nothing
Set db = Nothing

'setup for icon minimize to system tray
With IconData
    .cbSize = Len(IconData)
    .hIcon = Me.Icon
    .hwnd = Me.hwnd
    .szTip = "iBBS Server" & Chr(0)
    .uCallbackMessage = WM_MOUSEMOVE
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uID = vbNull
End With

'Setting up some starting variables
If blnServerStartState = True Then
    blnServerOn = True
    'blnServerPaused = False
    lngSock = 0
    lngFtp = 0
    intBuffer = 2048
    sckServer(lngSock).LocalPort = 1001
    'Start accepting connections
    sckServer(lngSock).Listen
    With txtClientOutput
        .SelColor = vbBlack
        .SelBold = 2
        .SelText = "iBBS Server Version: "
        .SelColor = vbGreen
        .SelText = App.Major & "." & App.Minor & " Revision " & App.Revision & vbCrLf
        .SelBold = 0
        .SelColor = vbBlue
        .SelBold = 2
        .SelText = TimeStamp & "Server Status: On" & vbCrLf
        
    End With
    
    lblusernm.Caption = "0"
Else
    blnServerOn = False
    'blnServerPaused = False
    lngSock = 0
    lngFtp = 0
    intBuffer = 2048
    sckServer(lngSock).LocalPort = 1001
    'Start accepting connections
    'sckServer(lngSock).Listen
    With txtClientOutput
        .SelColor = vbBlack
        .SelBold = 2
        .SelText = "iBBS Server Version: "
        .SelColor = vbGreen
        .SelText = App.Major & "." & App.Minor & " Revision " & App.Revision & vbCrLf
        .SelBold = 0
        .SelColor = vbRed
        .SelBold = 2
        .SelText = TimeStamp & "Server Status: Off" & vbCrLf
    End With
    
    lblusernm.Caption = "0"
End If
    

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim Msg As Long
Msg = X

'for double clicking icon in system tray to restore or show menu
If Msg = WM_LBUTTONDBLCLK Then
    Call mnuShow_Click
ElseIf Msg = WM_RBUTTONDOWN Then
    PopupMenu mnuPopup
End If



End Sub

Private Sub Form_Resize()

If Me.WindowState = 1 Then
    Call Shell_NotifyIcon(NIM_ADD, IconData)
    Me.Hide
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmMain = Nothing
Shell_NotifyIcon NIM_DELETE, IconData

End Sub

Private Sub sckServer_Close(Index As Integer)

Dim strUsers As String
Dim Z As Long

With txtClientOutput
    .SelColor = vbBlack
    .SelBold = 2
    .SelText = TimeStamp & "Disconnected:IP "
    .SelColor = vbRed
    .SelText = sckServer(Index).RemoteHostIP & vbCrLf
End With


'remove disconnected socket/user from collection
For Each User In colUsers
    If User.SockIndex = Index Then
        colUsers.Remove CStr(Index)
        Exit For
    End If
Next

sckServer(Index).Close
lngUsers = lngUsers - 1
If lngUsers < 0 Then lngUsers = 0
lblusernm.Caption = lngUsers

End Sub


Private Sub sckServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)

Dim X As Long

'Check to see if any open socks are being used
'and use unused ones for new connections to save memory
For X = 1 To sckServer.UBound
    If sckServer(X).State <> 7 Then
        sckServer(X).Close
        With txtClientOutput
            .SelColor = vbBlack
            .SelBold = 2
            .SelText = TimeStamp & "Incoming Request:IP "
            .SelColor = vbBlue
            .SelText = sckServer(Index).RemoteHostIP & " ID: " & requestID & vbCrLf
        End With
        sckServer(X).Accept requestID
        
        'send ok connect to client in order to receive login info
        sckServer(X).SendData "connect1||ok"
        GoTo exitconnect
    End If
Next

'If all open socks are being used, create a new one.
lngSock = lngSock + 1
Load sckServer(lngSock)
With txtClientOutput
    .SelColor = vbBlack
    .SelBold = 2
    .SelText = TimeStamp & "Incoming Request:IP "
    .SelColor = vbBlue
    .SelText = sckServer(Index).RemoteHostIP & " ID: " & requestID & vbCrLf
End With
'txtClientOutput.Text = txtClientOutput.Text & "Incoming Request:IP " & sckServer(Index).RemoteHostIP & " ID: " & requestID & vbCrLf
'txtClientOutput.Text = txtClientOutput.Text & vbCrLf & lngSock
sckServer(lngSock).Accept requestID
sckServer(X).SendData "connect1||ok"

exitconnect:

End Sub

Private Sub sckServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)

Dim lngInd As Long

'Get the incoming data from the client(s)
sckServer(Index).GetData strSendCode, vbStrin

'process sendcodes and send back results to client(s)
lngInd = Index
CheckSendCode strSendCode, lngInd


End Sub

Private Sub CheckSendCode(strCode As String, lngIndex As Long)
'Parse the inputed sendcode using split function. The sendcode is
'the first part of the incoming data. All incoming data comes in a
'style like "sendcode||handle||message" or something similar. But sendcode
'is always first. We will take that data and use the split function and pull
'out each piece marked by a "||" (double pipe) and do with it what is needed.
'This is so the server/client knows if the data is chat text, file data, etc.

Dim strHandle As String
Dim strPassword As String
Dim vntArray As Variant
Dim strText As String
Dim nItems As Integer
Dim n As Integer
Dim db As Database
Dim rs As Recordset

' split function will be used to parse items contained in a string,
' and delimitted by ||
' The Split function returns a variant array containing each parsed item
' as an element in the array

' use split function to parse it
vntArray = Split(strCode, "||")

' how many items were parsed?
nItems = UBound(vntArray)

'do a select case on the code type to determine clients request
'whether chat, im, mb, files, etc.. and take appropriate action
Select Case vntArray(0)
    Case "ibbslogin1"
        'login string
        strHandle = vntArray(1)
        strPassword = vntArray(2)
        With txtClientOutput
            .SelColor = vbBlack
            .SelBold = 2
            .SelText = TimeStamp & "Login attempt by: "
            .SelColor = vbBlue
            .SelText = strHandle & " " & vbCrLf
        End With
        
        'Open users database
        Set db = OpenDatabase(App.Path & "\ibbs.mdb")
        
        'Check if incoming IP is banned
        Set rs = db.OpenRecordset("Select * from bannedips where ipaddress = '" & sckServer(lngIndex).RemoteHostIP & "'")
        If rs.RecordCount <> 0 Then
            'this ip is banned
            sckServer(lngIndex).SendData "connect1||logonbanned"
            Set rs = Nothing
            Set db = Nothing
            Exit Sub
        End If
        
        'open user table in database and check user and password
        Set rs = db.OpenRecordset("Select * from users where handle = '" & strHandle & "' and password ='" & strPassword & "'")
        If rs.RecordCount <> 0 Then
            'login passed
            blnLoginGood = True
            'Let client know that logon is good and send thd client its
            'privileges, the server name, and the admin's name
            sckServer(lngIndex).SendData "connect1||logonyes||" & rs!Privileges & "||" & strServerName & "||" & strAdminName
            With txtClientOutput
                .SelColor = vbBlack
                .SelBold = 2
                .SelText = TimeStamp & "User Connected:IP "
                .SelColor = vbBlue
                .SelText = sckServer(Index).RemoteHostIP & " " & strID & vbCrLf
            End With

            lngUsers = lngUsers + 1
            lblusernm.Caption = lngUsers
            
            'Build and add user information to collection.
            'Make the SockIndex the key which will always be unique.
            With User
                .Handle = strHandle
                .SockIndex = lngIndex
                .IPAddress = sckServer(lngIndex).RemoteHostIP
                .UserPrivileges = rs!Privileges
                .ChatStatus = "OFF"
            End With
            
            colUsers.Add User, CStr(User.SockIndex)
            Set User = Nothing
                       
        Else
            'login failed
            blnLoginGood = False
            'Send bad logon string to client. Client will close connection
            sckServer(lngIndex).SendData "connect1||logonno"
            With txtClientOutput
                .SelColor = vbBlack
                .SelBold = 2
                .SelText = TimeStamp & "User Denied:IP "
                .SelColor = vbRed
                .SelText = sckServer(Index).RemoteHostIP & " " & strID & vbCrLf
            End With
            
        End If
                
        Set rs = Nothing
        Set db = Nothing
        
    Case "chatcode1"
        'chatroom string
        Dim strSendChatText As String
        
        'Create chat string to be sent to clients
        strSendChatText = "chatcode2||" & vntArray(2) & "||" & vntArray(1)
        With txtClientOutput
            .SelColor = vbBlack
            .SelBold = 2
            .SelText = TimeStamp & "Chat Text From "
            .SelColor = vbGreen
            .SelText = vntArray(2) & ": " & vntArray(1) & vbCrLf
        End With
        
        
        'loops through Users collection and sends
        'chat text only to those with ChatStatus set to "ON". Hopefully saving
        'more bandwidth and processing. Also send the text to the people in that
        'users channel
        For Each User In colUsers
            If User.ChatStatus = "ON" And User.ChatChannel = vntArray(3) Then
                sckServer(User.SockIndex).SendData strSendChatText
                DoEvents
            End If
        Next
    
    Case "chatstatus1"
        'change the users chat status
        If vntArray(1) = "ON" Then 'Chat for user is on
            For Each User In colUsers
                If User.Handle = vntArray(2) Then
                    User.ChatStatus = "ON"
                    User.ChatChannel = "Lobby"
                    'Send Channel list
                    sckServer(lngIndex).SendData "channellist1" & strChannels
                    Exit Sub
                End If
            Next
        Else 'Chat for user is off
            For Each User In colUsers
                If User.Handle = vntArray(2) Then
                    User.ChatStatus = "OFF"
                    User.ChatChannel = "Not in chat"
                    Exit Sub
                End If
            Next
        End If
                         
    Case "chatchannel1"
        'User is changing channel
        With txtClientOutput
            .SelColor = vbBlack
            .SelBold = 2
            .SelText = TimeStamp
            .SelColor = vbGreen
            .SelText = vntArray(2)
            .SelColor = vbBlack
            .SelText = " has changed to Chat Channel "
            .SelColor = vbBlue
            .SelText = vntArray(1) & vbCrLf
        End With
        
        For Each User In colUsers
            If User.Handle = vntArray(2) Then
                User.ChatChannel = vntArray(1)
                Exit Sub
            End If
        Next
        
    Case "filelistcode1"
        'file list request string
        Dim strSendList As String
        Dim strDir As String
        'do getfilelist function and return the list as a string
        With txtClientOutput
            .SelColor = vbBlack
            .SelBold = 2
            .SelText = TimeStamp & "Sending File list to "
            .SelColor = vbGreen
            .SelText = sckServer(lngIndex).RemoteHostIP & vbCrLf
        End With
        
        With colUsers(CStr(lngIndex))
            strDir = vntArray(1)
            'Compile list of files into string
            strSendList = GetFileList(strDir)
            'send list
            sckServer(lngIndex).SendData strSendList
        End With
        
    Case "getfilecode1"
        'Get file string
        Dim strFileName As String
        Dim BufferSize As Integer
        Dim lngFTPSock As Long
               
        strFileName = vntArray(1)
        With txtClientOutput
            .SelColor = vbBlack
            .SelBold = 2
            .SelText = TimeStamp & "Downloading "
            .SelColor = vbGreen
            .SelText = strFileName
            .SelColor = vbBlack
            .SelText = " from "
            .SelColor = vbBlue
            .SelText = sckServer(lngIndex).RemoteHostIP & vbCrLf
        End With
        
        lngBytesXfer = 0
        'Pass the users sock index and make an ftp winsock
        'connection to the user, this will cause problems with
        'firewalls, but I will deal with that later!
        lngFTPSock = MakeFTPConnection(lngIndex)
        fFile = FreeFile
    
        Open App.Path & "\files\" & strFileName For Binary Access Read As #fFile
    
        'loop through getfiledata getting each chunk of the file
        'and send it to the client until it is done
        Do Until FileLen(App.Path & "\files\" & strFileName) = lngBytesXfer
                
            DoEvents
            'Run the GetFileData function and send the chunk of
            'data
            strfiledata = GetFileData(FileLen(App.Path & "\files\" & strFileName), lngFTPSock, lngIndex)
            ftpSock1(lngFTPSock).SendData strfiledata
        Loop
               
        Close #fFile
        
        'Allow server to "catch up" before procceding to next file. Making
        'sure that last data chunk has been sent and file is closed.May fix
        'disconnecting that ocassionaly happens during file transfer
        Sleep 200
        
    Case "imcode1"
        'instant message string
        Dim strIMMessage As String
        Dim strIMTo As String
        Dim strIMFrom As String
        Dim lngItem As Long
        Dim strIMSendtoClient As String
        
        strIMMessage = vntArray(1)
        strIMTo = vntArray(2)
        strIMFrom = vntArray(3)
        
        'check collection of users and find sock of selected user
        For Each User In colUsers
            DoEvents
            If User.Handle = strIMTo Then
                lngItem = User.SockIndex
                'generate and send the message to the proper user
                strIMSendtoClient = "imcode2||" & strIMMessage & "||" & strIMFrom
                'send data to appropriate client
                sckServer(lngItem).SendData strIMSendtoClient
                Exit Sub
            End If
        Next
        
        'User not found. Send User not online message to client
        strIMSendtoClient = "imcode3||" & strIMTo
        sckServer(lngIndex).SendData strIMSendtoClient
        
    Case "userlist1"
        Dim channel As String
        
        'Online user list request
        strusersnew = "userlistcode2||"
        
        'Add users and what channel they are in,if any!
        For Each User In colUsers
            If User.ChatChannel = "" Then
                channel = "Not in chat"
            Else:
                channel = User.ChatChannel
            End If
            
            strusersnew = strusersnew & User.Handle & "@@" & channel & "||"
        Next
        
        'send new user list to requesting client
        sckServer(lngIndex).SendData strusersnew
                        
    Case "mbmessagescode1"
        'message board Get message list string
        'TODO
        'This is buggy at the moment. I will work on this,
        'when I get a better idea on how best to handle message boards
        Dim strSendMessageList As String
        Dim id As Long
        
        id = vntArray(1)
        With colUsers(CStr(lngIndex))
            'create list of messages
            strSendMessageList = GetMessageList(id)
            sckServer(lngIndex).SendData strSendMessageList
        End With
        
    Case "mbgroupscode1"
        'message board Get groups list string
        Dim strSendGroupList As String
        
        With colUsers(CStr(lngIndex))
            'create list of messages
            strSendGroupList = GetGroupList()
            sckServer(lngIndex).SendData strSendGroupList
        End With
    
    Case "mbgetmessagecode1"
        'get message board message
        Dim strSendMsg As String
        Dim lngid As Long
        
        lngid = vntArray(1)
        With colUsers(CStr(lngIndex))
            'create list of messages
            strSendMsg = GetMessage(lngid)
            sckServer(lngIndex).SendData strSendMsg
        End With
    
    Case "mbnewpostcode1"
        'add new message to message board
        
        Set db = OpenDatabase(App.Path & "\system\forum.mdb")
        Set rs = db.OpenRecordset("messages")
        
        'add the data to the database
        rs.AddNew
        rs!Handle = vntArray(1)
        rs!catid = vntArray(2)
        rs!subject = vntArray(3)
        rs!message = vntArray(4)
        rs!ip = sckServer(lngIndex).RemoteHostIP
        rs!Date = Now()
        
        rs.Update
        rs.Close
        db.Close
        
        
        Set rs = Nothing
        Set db = Nothing
        
        With txtClientOutput
            .SelColor = vbBlack
            .SelBold = 2
            .SelText = TimeStamp & "New Message Board Post by: "
            .SelColor = vbGreen
            .SelText = vntArray(1) & " Message ID: " & vntArray(2) & vbCrLf
        End With
        
    Case "news1"
        'Get latest Server news string
        Dim strNewsText As String
        Dim strLine As String
        
        'Get RTF news data and send it to requesting client
        filenum = FreeFile
        Open App.Path & "\system\news.rtf" For Input As #filenum
        Do Until EOF(filenum)
            Line Input #filenum, strLine
            strNewsText = strNewsText & strLine
        Loop
        Close #filenum
        'send news data to requesting client
        sckServer(lngIndex).SendData "news2||" & strNewsText
    
    Case "maillistcode1"
        'mail list string
        'This is for users personal email
        Dim strMailFor As String
        Dim strSendMailList As String
        
        strMailFor = vntArray(1)
        strSendMailList = "mailcode2||"
        
        Set db = OpenDatabase(App.Path & "\ibbs.mdb")
        
        'Get mail listing for user
        Set rs = db.OpenRecordset("Select m.subject, m.from, m.id from mail m inner join users u on u.id = m.touserid where u.handle ='" & strMailFor & "'")
        Do Until rs.EOF
            strSendMailList = strSendMailList & rs!subject & "@@" & rs!From & "@@" & rs!id & "||"
            rs.MoveNext
        Loop
        'Send listing to client
        sckServer(lngIndex).SendData strSendMailList
        
        rs.Close
        db.Close
        
        Set rs = Nothing
        Set db = Nothing
        
    
    Case "mailmessagecode1"
        Dim strSendMessage As String
    
        Set db = OpenDatabase(App.Path & "\ibbs.mdb")
        
        'Get mail message for user
        Set rs = db.OpenRecordset("Select mailmessage from mail where id=" & vntArray(1))
        
        'send selected mail message to client
        strSendMessage = "mailmessagecode2||" & rs!mailmessage
        sckServer(lngIndex).SendData strSendMessage
        
        rs.Close
        db.Close
        
        Set rs = Nothing
        Set db = Nothing
        
    Case "deletemailcode1"
        'delete mail string
        Set db = OpenDatabase(App.Path & "\ibbs.mdb")

        Set rs = db.OpenRecordset("select * from mail where id=" & vntArray(1))
        
        rs.Delete
        rs.Close
        db.Close
        
        Set rs = Nothing
        Set db = Nothing
        
    Case "newmailcode1"
        'new iBBS mail string
        Dim strTo As String
        Dim strFrom As String
        Dim strMailMsg As String
        Dim strSubject As String
        Dim lngid2 As Long
        
        strTo = vntArray(1)
        strMailMsg = vntArray(2)
        strFrom = vntArray(3)
        strSubject = vntArray(4)
        
        Set db = OpenDatabase(App.Path & "\ibbs.mdb")
        Set rs = db.OpenRecordset("Select id,handle from users where handle ='" & Trim(strTo) & "'")
        
        If rs.EOF Or rs.BOF Then
            'Invalid user
            sckServer(lngIndex).SendData "invalidmailcode1||"
            rs.Close
        Else
            'Valid user
            lngid2 = rs!id
            rs.Close
            Set rs = db.OpenRecordset("mail")
            rs.AddNew
            rs!touserid = lngid2
            rs!mailmessage = strMailMsg
            rs!From = strFrom
            rs!read = "No"
            rs!subject = strSubject
            rs.Update
            rs.Close
        End If
        
        db.Close
        Set rs = Nothing
        Set db = Nothing
        
    Case Else
        'invalid sendcode. Do nothing. Output details for review by Admin
        With txtClientOutput
            .SelColor = vbRed
            .SelBold = 2
            .SelText = TimeStamp & "Unknown send code. Possbile Hack Attempt. IP: " & sckServer(lngIndex).RemoteHostIP & "  String: " & strCode & vbCrLf
        End With
        
End Select

End Sub

Private Sub Timer1_Timer()

'Keep checking states so if there is a problem with a sock, it will close it
For X = 0 To sckServer.UBound
    With sckServer(X)
        If .State = 8 Or .State = 9 Then
            .Close
            Exit Sub
        End If
    End With
Next

End Sub

Private Sub mnuExit_Click()

Unload Me
End

End Sub

Private Sub mnuShow_Click()

Me.WindowState = vbNormal
Shell_NotifyIcon NIM_DELETE, IconData
Me.Show

End Sub

Private Function GetFileList(strWorkingDir As String) As String
Dim hFile As Long
Dim fname As String
Dim WFD As WIN32_FIND_DATA
Dim dirList As String

'Get the first file in the directory (it will usually return ".")
hFile = FindFirstFile(App.Path & "\files\" & strWorkingDir & "*.*" + Chr$(0), WFD)
    
'create string of files and their sizes to send
dirList = "filelistcode2||"
While FindNextFile(hFile, WFD)
    If Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> "." And Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) <> ".." Then
        dirList = dirList & WFD.nFileSizeLow & "\/" & Left$(WFD.cFileName, InStr(WFD.cFileName, vbNullChar) - 1) & "||"
    End If
    DoEvents
Wend

GetFileList = dirList
    
End Function

Private Function GetMessageList(lngid As Long) As String

Set db = OpenDatabase(App.Path & "\system\forum.mdb")
Set rs = db.OpenRecordset("Select * from Messages where catid =" & lngid)
strMBString = "mbmessagescode2||"
Do While Not rs.EOF
    strSubject = rs!subject
    strFrom = rs!Handle
    
    strID = rs!id
       
    
    strMBString = strMBString & strSubject & "\/" & strFrom & "\/" & strID & "||"
    rs.MoveNext

    
Loop

Set rs = Nothing
Set db = Nothing

GetMessageList = strMBString


End Function

Private Function MakeFTPConnection(ftpIndex As Long) As Long
Dim Y As Long

'I am using a seperate winsock control for files to prevent
'bottlenecking of data

'Check to see if any open ftp socks are being used
'and use unused ones for new connections to save memory
For Y = 1 To ftpSock1.UBound
    With ftpSock1(Y)
        If .State <> 7 Then
            .Close
            .RemoteHost = sckServer(ftpIndex).RemoteHostIP
            .RemotePort = "21"
            .Connect
            'Loop until it connects. This seems to have fixed the previous problem
            Do Until .State = 7
                DoEvents
            Loop
            MakeFTPConnection = Y
            GoTo exitftpconnect
        End If
    End With
Next

'If all open socks are being used, create a new one.
lngFtp = lngFtp + 1
Load ftpSock1(lngFtp)

With ftpSock1(lngFtp)
    .RemoteHost = sckServer(ftpIndex).RemoteHostIP
    .RemotePort = "21"
    .Connect
    Do Until .State = 7
        DoEvents
    Loop
End With

MakeFTPConnection = lngFtp

exitftpconnect:

End Function

Private Function GetFileData(ttlBytes As Long, lngsckFTP As Long, lngSock As Long)

Dim BlockSize As Integer
Dim DataToSend As String

BlockSize = intBuffer

'Determine the proper buffer size. If the remaining bytes to send is less than
'the block size(2048) then set block size to that value.
If BlockSize > (ttlBytes - lngBytesXfer) Then
    BlockSize = (ttlBytes - lngBytesXfer)
End If

DataToSend = Space$(BlockSize) 'allocate space to store data.
Get #fFile, , DataToSend 'get data chunk
'Assign the string of data to the function so it can be sent from the calling
'procedure
GetFileData = DataToSend

'Increment bytes sent
lngBytesXfer = lngBytesXfer + BlockSize
    
End Function

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button
    Case "Start/Stop Server"
        Command1_Click
    Case "Server Options"
        cmdOptions_Click
    Case "Close iBBS Server"
        Command2_Click
    Case "User's Online"
        Command3_Click
    Case "Send Admin Message"
        Command4_Click
    Case "Dump To Log"
        LogDump
End Select

End Sub

Private Function TimeStamp() As String

'Time stamp all output to the RTF box if the Time stamp output option is checked
If blnTStamp = True Then
    TimeStamp = "(" & Now() & ")  "
End If

End Function


Private Function GetGroupList() As String
Dim strGroupString As String

Set db = OpenDatabase(App.Path & "\system\forum.mdb")
Set rs = db.OpenRecordset("Select * from Groups")
strGroupString = "mbgroupscode2||"
'Build a string of groups.
Do Until rs.EOF
   
strGroupString = strGroupString & rs!category & "\/" & rs!id & "||"
rs.MoveNext

    
    
Loop

Set rs = Nothing
Set db = Nothing

GetGroupList = strGroupString


End Function

Private Function GetMessage(lngid As Long) As String

Set db = OpenDatabase(App.Path & "\system\forum.mdb")
Set rs = db.OpenRecordset("Select * from Messages where id =" & lngid)

If DateDiff("d", rs!Date, Date) > CInt(strDays) Then
    'message expired. Delete message and send back expire warning
    rs.Delete
    'need to compact databse too.
    strMsg = "mbmessageexpired2||" & lngid
Else:
    strMsg = "mbgetmessagecode2||" & rs!message
End If

rs.Close
db.Close

Set rs = Nothing
Set db = Nothing

GetMessage = strMsg

End Function
