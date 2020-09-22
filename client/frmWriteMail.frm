VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmWriteMail 
   Caption         =   "Compose New Mail"
   ClientHeight    =   4020
   ClientLeft      =   3840
   ClientTop       =   5475
   ClientWidth     =   5145
   Icon            =   "frmWriteMail.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4020
   ScaleWidth      =   5145
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1740
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtSubject 
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3480
      Width           =   1635
   End
   Begin VB.TextBox txtTo 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   60
      Width           =   3015
   End
   Begin RichTextLib.RichTextBox rtfMessage 
      Height          =   2475
      Left            =   0
      TabIndex        =   0
      Top             =   900
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   4366
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmWriteMail.frx":08CA
   End
   Begin VB.Label Label2 
      Caption         =   "Subject"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   480
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "To:"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   795
   End
End
Attribute VB_Name = "frmWriteMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim strNewMail As String
'Send mail string
If Trim(txtTo) = "" Then
    MsgBox ("You must enter the user name of the recipient!")
    Exit Sub
End If

If Trim(txtSubject) = "" Then
    MsgBox ("You must enter a subject!")
    Exit Sub
End If
    
If Trim(rtfMessage) = "" Then
    MsgBox ("Please enter a message!")
    Exit Sub
End If

strNewMail = "newmailcode1||" & txtTo.Text & "||" & rtfMessage.Text & "||" & strHandle & "||" & txtSubject.Text

MDIForm1.sckClient.SendData strNewMail

Unload frmWriteMail

End Sub

Private Sub Command2_Click()

Unload frmWriteMail

End Sub


Private Sub Form_Unload(Cancel As Integer)

Set frmWriteMail = Nothing

End Sub
