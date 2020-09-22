VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connecting to St Louis Weather"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   4080
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   1680
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   1680
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1080
      Top             =   1080
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   " Logging Into St Louis Weather Servers."
      Top             =   120
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1080
      Width           =   150
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Text2.Text = " Logging Into St Louis Weather Servers.."
Timer1.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub Timer2_Timer()
Text2.Text = " Logging Into St Louis Weather Servers..."
Timer2.Enabled = False
Timer3.Enabled = True
End Sub

Private Sub Timer3_Timer()
Text2.Text = " Logging Into St Louis Weather Servers...."
Timer3.Enabled = False
Timer5.Enabled = True
End Sub

Private Sub Timer5_Timer()
Text2.Text = " Logging Into St Louis Weather Servers."
Timer5.Enabled = False
Timer1.Enabled = True
End Sub
