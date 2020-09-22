VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form chatform 
   Caption         =   "Chat Room"
   ClientHeight    =   8115
   ClientLeft      =   5475
   ClientTop       =   2385
   ClientWidth     =   6600
   Icon            =   "chatform.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   6600
   Begin VB.CommandButton Command2 
      Caption         =   "Change Channel"
      Height          =   315
      Left            =   2100
      TabIndex        =   5
      Top             =   0
      Width           =   1455
   End
   Begin VB.ComboBox cbChannel 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Text            =   "Lobby"
      Top             =   0
      Width           =   2055
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   6975
      Left            =   0
      TabIndex        =   3
      Top             =   420
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   12303
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"chatform.frx":08CA
   End
   Begin VB.ListBox List1 
      Height          =   6885
      Left            =   4740
      TabIndex        =   2
      Top             =   420
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send "
      Height          =   495
      Left            =   5100
      TabIndex        =   1
      Top             =   7560
      Width           =   1275
   End
   Begin VB.TextBox Text3 
      Height          =   435
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   7560
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Users In Channel"
      Height          =   255
      Left            =   4740
      TabIndex        =   6
      Top             =   180
      Width           =   1755
   End
   Begin VB.Line Line2 
      X1              =   6480
      X2              =   60
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line1 
      X1              =   4680
      X2              =   4680
      Y1              =   240
      Y2              =   7320
   End
End
Attribute VB_Name = "chatform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()

'Send chat text
MDIForm1.sckClient.SendData "chatcode1||" & Text3.Text & "||" & strHandle & "||" & strChannel

Text3.Text = ""

End Sub

Private Sub Command2_Click()

'Send code to server to change chat channel
strChannel = cbChannel.Text
MDIForm1.sckClient.SendData "chatchannel1||" & strChannel & "||" & strHandle
With Text1
    .SelColor = vbRed
    .SelText = "You are now joining Channel: " & strChannel & vbCrLf
End With

End Sub

Private Sub Form_Load()
strChannel = "Lobby"
With Text1
    .SelColor = vbRed
    .SelText = "You have joined Channel: Lobby" & vbCrLf
End With

strChatFormState = "Open"
'Send chat ON status to server
MDIForm1.sckClient.SendData "chatstatus1||ON||" & strHandle


End Sub

Private Sub Form_Unload(Cancel As Integer)

'Set ths variable so that I know not to print
'incoming chat text. I may re write it so that
'the server doesnt send it either.
strChatFormState = "Closed"

'send chat off status to server
MDIForm1.sckClient.SendData "chatstatus1||OFF||" & strHandle

Set chatform = Nothing

End Sub

Private Sub Text3_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Command1_Click
End If

End Sub
