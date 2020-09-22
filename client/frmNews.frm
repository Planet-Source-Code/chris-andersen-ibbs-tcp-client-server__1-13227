VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmNews 
   Caption         =   "Latest News"
   ClientHeight    =   2655
   ClientLeft      =   2220
   ClientTop       =   3675
   ClientWidth     =   9135
   Icon            =   "frmNews.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2655
   ScaleWidth      =   9135
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmNews.frx":08CA
   End
End
Attribute VB_Name = "frmNews"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)

Set frmNews = Nothing

End Sub
