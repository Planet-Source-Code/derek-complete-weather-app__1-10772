VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configure St Louis Dopplar Radar Tool"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   4455
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox DDEText 
      Height          =   375
      Left            =   1080
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   240
      Top             =   2040
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   1440
      Top             =   3480
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   9
      Text            =   "Manu"
      ToolTipText     =   "Current Status"
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Turn Conditions Off"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      ToolTipText     =   "Turn Off Auto Condition Download"
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Auto Conditions Timer"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      ToolTipText     =   "Download and Save Conditions Every 60 Sec"
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Turn Auto Zoom Off"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "Auto Zoom Off"
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   5
      Text            =   "Manu"
      ToolTipText     =   "Current Status"
      Top             =   3240
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Auto Zoom Download"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Zoom to City on Download"
      Top             =   3240
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Auto Download Off"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      ToolTipText     =   "Auto Download Off"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3600
      MaxLength       =   5
      TabIndex        =   2
      Text            =   "60000"
      ToolTipText     =   "Time in Ms, 60000 = 60s"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download Timer On"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Auto Download Radar"
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4800
      Width           =   375
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo er:
Text1.SetFocus
Form1.Timer1.Interval = Text2.Text
Form1.Timer1.Enabled = True
Form4.Hide
Exit Sub
er:
MsgBox "Must Enter a Number Between 0 and 60000", vbSystemModal, "Internet Radar Timer Refresh"
Form1.Timer1.Enabled = False
End Sub

Private Sub Command2_Click()
Text1.SetFocus
Form1.Timer1.Enabled = False
Form4.Hide
Exit Sub
End Sub

Private Sub Command3_Click()
Text1.SetFocus
Text1.Text = "y"
Text3.Text = "Auto"
Form4.Hide
Exit Sub
End Sub

Private Sub Command4_Click()
Text1.SetFocus
Text1.Text = "n"
Text3.Text = "Manu"
Form4.Hide
Exit Sub
End Sub

Private Sub Command5_Click()
Text1.SetFocus
Text4.Text = "Manu"
Form4.Hide
Timer1.Enabled = False
Timer2.Enabled = False
Exit Sub
End Sub

Private Sub Command6_Click()
Text1.SetFocus
Text4.Text = "Auto"
Form4.Hide
Timer1.Enabled = True
Exit Sub
End Sub

Private Sub Command7_Click()

Timer3.Enabled = True
Form4.Hide
Exit Sub
End Sub

Private Sub Command8_Click()
Timer3.Enabled = False
Form4.Hide
Exit Sub
End Sub



Private Sub Timer1_Timer()
Form5.Command6.Value = True
Timer2.Enabled = True
Exit Sub
End Sub

Private Sub Timer2_Timer()
If Form5.Command6.Enabled = True Then
Form5.Command4.Value = True
Timer2.Enabled = False
Else
Timer2.Enabled = True
End If
End Sub



