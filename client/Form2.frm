VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmIM 
   Caption         =   "Instant Message"
   ClientHeight    =   3945
   ClientLeft      =   8250
   ClientTop       =   5505
   ClientWidth     =   5970
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3945
   ScaleWidth      =   5970
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3135
      Left            =   0
      TabIndex        =   2
      Top             =   60
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5530
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form2.frx":08CA
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   4920
      TabIndex        =   1
      Top             =   3300
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   555
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3300
      Width           =   4635
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strIM As String

Private Sub Command1_Click()

'RC4 encode message
strIM = "imcode1||" & RC4(Text2.Text, key) & "||" & Me.Caption & "||" & strHandle
'strIM = "imcode1||" & Text2.Text & "||" & Me.Caption & "||" & strHandle

MDIForm1.sckClient.SendData strIM
With Text1
    .SelColor = vbGreen
    .SelBold = 2
    .SelText = "<" & strHandle & ">"
    .SelColor = vbBlack
    .SelText = Text2.Text & vbCrLf
End With
'Clear out the data to send text box

Text2.SetFocus

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Command1_Click
End If

End Sub

