VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Your Current Weather Watches / Warnings"
   ClientHeight    =   570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "Form10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text6 
      Height          =   1575
      Left            =   2880
      TabIndex        =   8
      Text            =   "Text6"
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scan County"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Scan For Warnings"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Text            =   "Text5"
      Top             =   120
      Width           =   150
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      ToolTipText     =   "Scan For Warnings"
      Top             =   550
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "None"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      ToolTipText     =   "Scan For Warnings"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Text4 
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
      Height          =   380
      Left            =   120
      TabIndex        =   3
      Text            =   " No Counties Scanned For Active Statments..."
      Top             =   550
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   1430
      TabIndex        =   2
      Text            =   "ST. LOUIS"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Height          =   360
      Left            =   2985
      TabIndex        =   1
      Text            =   "MISSOURI"
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Text            =   "Text3"
      Top             =   1080
      Width           =   2175
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1080
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327681
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Text3.SetFocus
 Call yyyy
End Sub

Private Sub Command2_Click()
Text3.SetFocus
Command2.Enabled = False
Command3.Enabled = False
Form10.Height = 945
Text4.Visible = False: Command2.Visible = False
End Sub

Private Sub Command3_Click()
Text3.SetFocus
Command2.Enabled = True
Form10.Height = 1365
Text4.Visible = True: Command2.Visible = True
Exit Sub
End Sub

Private Sub Form_Load()
Dim ReturnStr As String
Text3.Text = Inet1.OpenURL("http://iwin.nws.noaa.gov/iwin/us/thunderstorm.html", icString)
ReturnStr = Inet1.GetChunk(2048, icString)

Do While Len(ReturnStr) <> 0


    DoEvents
        Text3.Text = Text3.Text & ReturnStr
        ReturnStr = Inet1.GetChunk(2048, icString)
    Loop
    
    Dim ReturnStr1 As String
Text6.Text = Inet1.OpenURL("http://iwin.nws.noaa.gov/iwin/us/tornado.html", icString)
ReturnStr1 = Inet1.GetChunk(2048, icString)

Do While Len(ReturnStr1) <> 0


    DoEvents
        Text6.Text = Text6.Text & ReturnStr
        ReturnStr1 = Inet1.GetChunk(2048, icString)
    Loop
    
    
    Call yyyy
    
 Command1.Enabled = True
Form7.Hide
End Sub

Sub xxxx()
Dim x
x = InStr(1, Text3.Text, Text1.Text)
If x = 0 Then
Command3.Enabled = False
Command3.Caption = "None"
Text4.Text = " No Warnings Found For " + Text1.Text
Command3.Enabled = True
Command3.Enabled = False
Call zzzz
Else
Command3.Enabled = True
Command3.Caption = "Read"
Text4.Text = " Thunderstorm Warning Active : " + Text1.Text
End If
End Sub
Sub yyyy()
Dim x
x = InStr(1, Text6.Text, Text1.Text)
If x = 0 Then
Command3.Enabled = False
Command3.Caption = "None"
Text4.Text = " No Warnings Found For " + Text1.Text
Command3.Enabled = True
Command3.Enabled = False
Call xxxx
Else
Command3.Enabled = True
Command3.Caption = "Read"
Text4.Text = " Tornado Warning Active : " + Text1.Text
End If
End Sub


Sub zzzz()
Dim x
x = InStr(1, Text6.Text, Text2.Text)
If x = 0 Then
Command3.Enabled = False
Command3.Caption = "None"
Text4.Text = " No Warnings Found For " + Text2.Text
Command3.Enabled = True
Command3.Enabled = False
Call aaaa
Else
Command3.Enabled = True
Command3.Caption = "Read"
Text4.Text = " Tornado Warnings In " + Text2.Text
End If
End Sub

Sub aaaa()
Dim x
x = InStr(1, Text3.Text, Text2.Text)
If x = 0 Then
Command3.Enabled = False
Command3.Caption = "None"
Text4.Text = " No Warnings Found For " + Text2.Text
Command3.Enabled = True
Command3.Enabled = False
Exit Sub
Else
Command3.Enabled = True
Command3.Caption = "Read"
Text4.Text = " Tunderstorm Warnings In " + Text2.Text
End If
End Sub
