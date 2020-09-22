VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pan Radar In Any Direction"
   ClientHeight    =   975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   3615
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "South"
      Height          =   735
      Left            =   2640
      Picture         =   "Form3.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "East"
      Height          =   735
      Left            =   1800
      Picture         =   "Form3.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "North"
      Height          =   735
      Left            =   960
      Picture         =   "Form3.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "West"
      Height          =   735
      Left            =   120
      Picture         =   "Form3.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1560
      Width           =   150
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.SetFocus
With Form1.Image1
.Left = .Left + 100
End With
Exit Sub
End Sub

Private Sub Command2_Click()
Text1.SetFocus
With Form1.Image1
.Top = .Top + 100
End With
Exit Sub
End Sub

Private Sub Command3_Click()
Text1.SetFocus
With Form1.Image1
.Left = .Left - 100
End With
Exit Sub
End Sub

Private Sub Command4_Click()
Text1.SetFocus
With Form1.Image1
.Top = .Top - 100
End With
Exit Sub
End Sub
