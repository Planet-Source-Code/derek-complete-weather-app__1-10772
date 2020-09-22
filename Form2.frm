VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Your County"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   3015
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "Form2.frx":030A
      Left            =   120
      List            =   "Form2.frx":032C
      TabIndex        =   4
      Text            =   "Select a Current City or County"
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   120
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit Without Zooming"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   500
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Zoom City"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   500
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1680
      Width           =   150
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
If Form2.Visible = True Then Text1.SetFocus
If Combo1.ListIndex = 0 Then
Form1.Image1.Stretch = False
Form1.Image1.Width = Form1.Image1.Width * 200 / 100
Form1.Image1.Height = Form1.Image1.Height * 200 / 100
Form1.Image1.Left = -6000
Form1.Image1.Top = -4720
Form1.Image1.Stretch = True
End If
If Combo1.ListIndex = 1 Then
Form1.Image1.Stretch = False
Form1.Image1.Width = Form1.Image1.Width * 200 / 100
Form1.Image1.Height = Form1.Image1.Height * 200 / 100
Form1.Image1.Left = -7000
Form1.Image1.Top = -5720
Form1.Image1.Stretch = True
End If
If Combo1.ListIndex = 2 Then
Form1.Image1.Stretch = False
Form1.Image1.Width = Form1.Image1.Width * 200 / 100
Form1.Image1.Height = Form1.Image1.Height * 200 / 100
Form1.Image1.Left = -8100
Form1.Image1.Top = -5720
Form1.Image1.Stretch = True
End If
If Combo1.ListIndex = 3 Then
Form1.Image1.Stretch = False
Form1.Image1.Width = Form1.Image1.Width * 200 / 100
Form1.Image1.Height = Form1.Image1.Height * 200 / 100
Form1.Image1.Left = -3100
Form1.Image1.Top = -4720
Form1.Image1.Stretch = True
End If
If Combo1.ListIndex = 4 Then
Form1.Image1.Stretch = False
Form1.Image1.Width = Form1.Image1.Width * 200 / 100
Form1.Image1.Height = Form1.Image1.Height * 200 / 100
Form1.Image1.Left = -4820
Form1.Image1.Top = -2720
Form1.Image1.Stretch = True
End If
If Combo1.ListIndex = 5 Then
Form1.Image1.Stretch = False
Form1.Image1.Width = Form1.Image1.Width * 200 / 100
Form1.Image1.Height = Form1.Image1.Height * 200 / 100
Form1.Image1.Left = -3500
Form1.Image1.Top = -1500
Form1.Image1.Stretch = True
End If
If Combo1.ListIndex = 6 Then
Form1.Image1.Stretch = False
Form1.Image1.Width = Form1.Image1.Width * 200 / 100
Form1.Image1.Height = Form1.Image1.Height * 200 / 100
Form1.Image1.Left = -6500
Form1.Image1.Top = -8500
Form1.Image1.Stretch = True
End If
If Combo1.ListIndex = 7 Then
Form1.Image1.Stretch = False
Form1.Image1.Width = Form1.Image1.Width * 200 / 100
Form1.Image1.Height = Form1.Image1.Height * 200 / 100
Form1.Image1.Left = -6500
Form1.Image1.Top = -8500
Form1.Image1.Stretch = True
End If
If Combo1.ListIndex = 8 Then
Form1.Image1.Stretch = False
Form1.Image1.Width = Form1.Image1.Width * 200 / 100
Form1.Image1.Height = Form1.Image1.Height * 200 / 100
Form1.Image1.Left = -7320
Form1.Image1.Top = -4720
Form1.Image1.Stretch = True
End If
If Combo1.ListIndex = 9 Then
Form1.Image1.Stretch = False
Form1.Image1.Width = Form1.Image1.Width * 200 / 100
Form1.Image1.Height = Form1.Image1.Height * 200 / 100
Form1.Image1.Left = -2000
Form1.Image1.Top = -4720
Form1.Image1.Stretch = True
End If
Form1.Image1.Stretch = True
Form2.Visible = False
Form2.Hide
Form1.Show
Exit Sub
End Sub

Private Sub Command2_Click()
Text1.SetFocus
Form2.Hide
Form2.Visible = False
Form1.Show
Exit Sub
End Sub
