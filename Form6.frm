VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Send Message"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2505
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2505
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   " Remote Computer Ip"
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form6.frx":030A
      Top             =   450
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send Instant Message to Ip"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Send Instant Message"
      Top             =   1230
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   150
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
With Form1
.Text1.SetFocus
.Winsock2.Close
.Winsock2.RemoteHost = Form6.Text1.Text
.Winsock2.RemotePort = 22098
.Winsock2.Connect
.Timer2.Enabled = True
End With
Exit Sub
End Sub

Private Sub Command2_Click()
Text1.SetFocus
Form6.Hide
End Sub

