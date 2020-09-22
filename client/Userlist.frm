VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Userlist 
   Caption         =   "Users Online Now"
   ClientHeight    =   8355
   ClientLeft      =   9255
   ClientTop       =   2145
   ClientWidth     =   3270
   Icon            =   "Userlist.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   3270
   Begin MSComctlLib.ListView userlist1 
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   13996
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
         Text            =   "User"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "In Chat"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Double Click User To IM"
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "Userlist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Me.Height = 8760
Me.Width = 3390

End Sub

Private Sub userlist1_DblClick()

Load IMForm(IMNumber)
With IMForm(IMNumber)
    .Height = 4350
    .Width = 6090
    .Caption = userlist1.SelectedItem.Text
    .Show
End With
IMNumber = IMNumber + 1

End Sub
