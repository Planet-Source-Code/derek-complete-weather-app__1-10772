VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set New Dopplar Radar Location Download"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   4920
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   420
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "Set New Download Path"
      Top             =   120
      Width           =   615
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
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Text            =   "http://www.kdsk.com/radar_data/max40.gif"
      Top             =   120
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
If Text2.Text = "http://www.kdsk.com/radar_data/max40.gif" Then Text1.Text = "ok"
If Text2.Text = "Http://www.ksdk.com/radar_data/max40.gif" Then Text1.Text = "ok"
If Text2.Text = "http://www.ksdk.com/radar_data/max40.gif" Then Text1.Text = "ok"
If Text2.Text = "www.ksdk.com/radar_data/max40.gif" Then Text1.Text = "ok"
If Text2.Text = "www.Ksdk.com/Radar_Data/Max40.gif" Then Text1.Text = "ok"
If Text2.Text = "http://www.Ksdk.com/Radar_Data/Max40.gif" Then Text1.Text = "ok"
If Text2.Text = "Http://www.Ksdk.com/Radar_Data/Max40.gif" Then Text1.Text = "ok"
Form1.Text3.Text = Text2.Text
If Text1.Text = "ok" Then
Form1.Command3.Enabled = True
Form5.Command3.Enabled = True
Else
Form1.Command3.Enabled = False
Form5.Command3.Enabled = False
End If
Form8.Hide
End Sub

Private Sub Form_Load()
Text1.Text = "no"
End Sub
