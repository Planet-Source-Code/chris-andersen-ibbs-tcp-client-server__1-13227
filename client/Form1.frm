VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLogin 
   Caption         =   "iBBS Client Logon"
   ClientHeight    =   2340
   ClientLeft      =   5340
   ClientTop       =   5370
   ClientWidth     =   5730
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   5730
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   780
      Left            =   2580
      TabIndex        =   6
      Top             =   1560
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   1376
      ButtonWidth     =   1508
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Logon"
            Object.ToolTipText     =   "Logon to server"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Object.ToolTipText     =   "Close iBBS Client"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   900
      Top             =   1740
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   50
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1740
      TabIndex        =   5
      Top             =   600
      Width           =   3435
   End
   Begin VB.TextBox Text2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1740
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   3435
   End
   Begin VB.TextBox txtIP 
      Height          =   375
      Left            =   1740
      TabIndex        =   1
      Top             =   120
      Width           =   3435
   End
   Begin VB.Label Label3 
      Caption         =   "Handle"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   660
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Address"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'On Error GoTo connecterror
strHost = Trim(txtIP.Text)
strUser = Text3.Text
strPass = Text2.Text

Load MDIForm1
MDIForm1.Hide

Exit Sub

connecterror:
'If Err.Number = "40006" Then
'    MsgBox ("Server not responding!")
'End If

End Sub

Private Sub Form_Load()
Dim vntCMDLine As Variant


If Command = "" Then
    txtIP.Text = "127.0.0.1"
    strHost = Trim(txtIP.Text)
    
Else
    vntCMDLine = Split(Command, ":")
    txtIP.Text = Right(vntCMDLine(1), Len(vntCMDLine(1)) - 2)
    Text3.Text = vntCMDLine(2)
    Text2.Text = Left(vntCMDLine(3), Len(vntCMDLine(3)) - 1)
    
    Command1_Click
End If

    

'MDIForm1.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)

'End

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button
    Case "Logon"
        Command1_Click
    Case "Close"
        End
End Select

End Sub
