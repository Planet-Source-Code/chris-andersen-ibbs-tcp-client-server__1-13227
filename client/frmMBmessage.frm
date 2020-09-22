VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMBmessage 
   Caption         =   "New Message Board Post"
   ClientHeight    =   4305
   ClientLeft      =   5445
   ClientTop       =   4950
   ClientWidth     =   6585
   Icon            =   "frmMBmessage.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4305
   ScaleWidth      =   6585
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1500
      TabIndex        =   6
      Top             =   60
      Width           =   1335
   End
   Begin VB.TextBox txtSubject 
      Height          =   345
      Left            =   1560
      TabIndex        =   5
      Top             =   900
      Width           =   4875
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Post"
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rtfMessage 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   1260
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5318
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmMBmessage.frx":08CA
   End
   Begin VB.Label Label3 
      Caption         =   "Subject:"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label Label2 
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Posting to Group:"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmMBmessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Trim(txtSubject.Text) = "" Then
    MsgBox ("You must enter a subject!")
    Exit Sub
End If

If Trim(rtfMessage.Text) = "" Then
    MsgBox ("Please enter a message!")
    Exit Sub
End If

MDIForm1.sckClient.SendData "mbnewpostcode1||" & strHandle & "||" & lngMBGroup & "||" & txtSubject.Text & "||" & rtfMessage.Text

Unload frmMBmessage

End Sub

Private Sub Command2_Click()
Unload frmMBmessage

End Sub

Private Sub Form_Load()

Label2.Caption = strMBGroup

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set frmMBmessage = Nothing

End Sub
