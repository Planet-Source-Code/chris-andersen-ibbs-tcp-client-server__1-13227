VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fileform 
   Caption         =   "File Section"
   ClientHeight    =   7740
   ClientLeft      =   8145
   ClientTop       =   2355
   ClientWidth     =   4050
   Icon            =   "fileform.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   4050
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   582
      ButtonWidth     =   3043
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Download Tagged"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4560
      Top             =   6060
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":1A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":2358
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":2C32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":350C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":3DE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":46C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fileform.frx":4F9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7080
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSWinsockLib.Winsock ftpclient 
      Left            =   5280
      Top             =   6180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   11033
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblStatus 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Double Click File To Download Or Tag Files and Click Download Tagged"
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4035
   End
End
Attribute VB_Name = "fileform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file As ListItem

Private Sub Command1_Click()
'Start loop to download checked files

For Each file In ListView1.ListItems
    If file.Checked = True Then
        strStatus = "Downloading"
        lblStatus.Caption = strStatus
        MDIForm1.Timer1.Interval = 0
        MDIForm1.sckClient.SendData "getfilecode1||" & file
        fFile = FreeFile
        MDIForm1.Timer1.Interval = 10000
        
        strFileLen = file.SubItems(1)
        With ProgressBar1
            .Value = 0
            .Max = strFileLen
        End With
        
        Open App.Path & "\dl\" & file For Binary Access Write As #fFile
                
    
        Do Until strStatus = "Ready"
            DoEvents
        Loop
        file.Checked = False
    End If
Next

End Sub

Private Sub ftpClient_ConnectionRequest(ByVal requestID As Long)

ftpclient.Close
ftpclient.Accept requestID

End Sub

Private Sub ftpClient_DataArrival(ByVal bytesTotal As Long)

Dim data As String

ftpclient.GetData data

'Increment progress bar to show files download status
ProgressBar1.Value = ProgressBar1.Value + bytesTotal
'Insert incoming data into the open file
Put #fFile, , data

'If file's total length is equal to the current byte
'position, close the file and get ready to download the
'next file
If strFileLen = Loc(fFile) Then
    Close #fFile
    ftpclient.Close
    'A pause for good measure
    Sleep 200
    ftpclient.Listen
    strFileLen = 0
    strStatus = "Ready"
    lblStatus.Caption = strStatus
End If

End Sub

Private Sub Form_Load()

ftpclient.LocalPort = "21"
ftpclient.Listen
MDIForm1.Timer1.Interval = 0
MDIForm1.sckClient.SendData "filelistcode1||"
MDIForm1.Timer1.Interval = 10000

End Sub

Private Sub ListView1_DblClick()

MDIForm1.sckClient.SendData "getfilecode1||" & fileform.ListView1.SelectedItem

fFile = FreeFile

strFileLen = fileform.ListView1.SelectedItem.SubItems(1)
With ProgressBar1
    .Value = 0
    .Max = strFileLen
End With

Open App.Path & "\dl\" & fileform.ListView1.SelectedItem For Binary Access Write As #fFile


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button
    Case "Download Tagged"
        Command1_Click
End Select

End Sub
