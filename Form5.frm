VERSION 5.00
Object = "{CAC6CA47-B177-48C9-B7D0-AEBE8F9B3F9B}#1.0#0"; "WEATHERWATCH.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advanced Radar Options"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3495
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Text            =   "Current Weather Condition - Waiting for File"
      Top             =   2750
      Width           =   3255
   End
   Begin WeatherWatcher.WeatherWatch WeatherWatch1 
      Left            =   360
      Top             =   3240
      _ExtentX        =   1296
      _ExtentY        =   873
   End
   Begin VB.TextBox Text14 
      BackColor       =   &H00E0E0E0&
      Height          =   1605
      Left            =   -240
      TabIndex        =   16
      Top             =   4080
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save as File"
      Height          =   855
      Left            =   2280
      Picture         =   "Form5.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Save Condition to \conditions.txt"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Radar Info"
      Height          =   855
      Left            =   1200
      Picture         =   "Form5.frx":0775
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Download Radar Info"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Conditions"
      Height          =   855
      Left            =   120
      Picture         =   "Form5.frx":0BB7
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Return Conditions"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Waiting"
      ToolTipText     =   "Barometer"
      Top             =   2055
      Width           =   735
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Waiting"
      ToolTipText     =   "Sunrise"
      Top             =   2055
      Width           =   735
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Waiting"
      ToolTipText     =   "Sunset"
      Top             =   2385
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Waiting"
      ToolTipText     =   "Visibility"
      Top             =   2385
      Width           =   735
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Waiting"
      ToolTipText     =   "Humidity"
      Top             =   2055
      Width           =   735
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Waiting"
      ToolTipText     =   "Dewpoint"
      Top             =   2055
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Waiting"
      ToolTipText     =   "Wind"
      Top             =   2385
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Waiting"
      ToolTipText     =   "Temp"
      Top             =   2385
      Width           =   735
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Current Time Now - Unknown"
      Top             =   1410
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Raw File Size - Unknow Size Limit"
      Top             =   3360
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Radar Mode - Uknown Mode Type"
      Top             =   1725
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Radar Updated - Unknown Time"
      Top             =   1080
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   4440
      Width           =   1575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Sub

Private Sub Command2_Click()

Exit Sub
End Sub

Private Sub Command3_Click()
Text1.SetFocus
If Text1.Text = "yes" Then
Text5.Text = "Current Time Now - " & Time
Text1.Text = Time
Text4.Text = "Raw File Size - " & Text14.Text & " Bytes"
If Form1.Text2.Text = "" Then Exit Sub:
Text2.Text = "Raw Updated - " + Form1.Text2.Text
If Left(Text14.Text, 2) >= 35 Then
Text3.Text = "Radar Mode - Precipitation Mode"
Else
Text3.Text = "Radar Mode - Clean Air Mode"
End If
Else
Exit Sub
End If
End Sub



Private Sub Command4_Click()
On Error Resume Next
Open App.path + "\conditions.txt" For Output As #1
Print #1, "Current Temperature Reported : " + Text6.Text; ""
Print #1, "Wind Speed Detected : " + Text7.Text
Print #1, "Current Dewpoint is : " + Text8.Text
Print #1, "Humidity Reported is : " + Text9.Text
Print #1, "Current Visibility is : " + Text10.Text
Print #1, "Your Sun Sets at : " + Text11.Text
Print #1, "Your Sun Rises at : " + Text12.Text
Print #1, "Barometer Reading is : " + Text13.Text
Print #1, Text14.Text
Close #1
Command4.Enabled = True
Form5.Hide
End Sub

Private Sub Command6_Click()
Text1.SetFocus
Command6.Enabled = False
Call condition
gett
End Sub

Sub gett()
On Error GoTo ex:
Command6.Enabled = False
WeatherWatch1.City = Form9.Text4.Text
WeatherWatch1.State = Form9.Text2.Text
WeatherWatch1.Connect
Text6.Text = WeatherWatch1.Temperature
Text7.Text = WeatherWatch1.WindSpeed
Text8.Text = WeatherWatch1.DewPoint
Text9.Text = WeatherWatch1.Humidity
Text13.Text = WeatherWatch1.Barometer
Text12.Text = WeatherWatch1.Sunrise
Text11.Text = WeatherWatch1.Sunset
Text10.Text = WeatherWatch1.Visibility
Command6.Enabled = True
Exit Sub
ex:
Command6.Enabled = True
End Sub

Sub condition()
On Error Resume Next
Dim Text As String
Dim Search As String
Dim Spot As Integer
Dim Spot2 As Integer
Search = "<FONT FACE=""Arial, Helvetica, Chicago, Sans Serif"" SIZE=3><B>"
Text = Form1.sock.OpenURL("http://www.weather.com/weather/us/zips/" & Form9.Text3.Text & ".html")
Spot = InStr(1, Text, Search) + Len(Search)
Spot2 = InStr(Spot, Text, "</B>")
returned = Mid$(Text, Spot, Spot2 - Spot)
Text15.Text = "Current Weather Condition - " + returned
End Sub
