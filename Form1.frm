VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "St Louis Dopplar Radar Tool 1.0 - Email Derekgregg@mail.ru"
   ClientHeight    =   4185
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6615
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet sock 
      Left            =   1920
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327681
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   1560
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1800
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "Message From Remote Software"
      Top             =   5160
      Width           =   5295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Send Msg"
      Height          =   855
      Left            =   120
      Picture         =   "Form1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Send Message to Remote Software"
      Top             =   4560
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1800
      Top             =   6000
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   2520
      TabIndex        =   11
      Text            =   " Type Your Message And Press The Message Button"
      ToolTipText     =   "Text to Send"
      Top             =   4800
      Width           =   3975
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Text            =   "Remote Address"
      ToolTipText     =   "Ip Address"
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Text            =   "http://www.ksdk.com/radar_data/max40.gif"
      Top             =   5880
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   6000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3840
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   5880
      Width           =   150
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   1200
      ScaleHeight     =   4185
      ScaleWidth      =   5265
      TabIndex        =   6
      ToolTipText     =   "Advanced Settings"
      Top             =   120
      Width           =   5295
      Begin VB.Image Image1 
         Height          =   4935
         Left            =   -1320
         Top             =   -1320
         Width           =   6015
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Advanced"
      Height          =   855
      Left            =   120
      Picture         =   "Form1.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Advanced Radar Settings"
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pan Radar"
      Height          =   855
      Left            =   120
      Picture         =   "Form1.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Pan Map In Any Direction"
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Zoom Map"
      Height          =   855
      Left            =   120
      Picture         =   "Form1.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Zoom in Towards a City"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Download"
      Height          =   855
      Left            =   120
      Picture         =   "Form1.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Download Needed Files"
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Configure"
      Height          =   855
      Left            =   120
      MaskColor       =   &H00C0E0FF&
      Picture         =   "Form1.frx":18A0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Configure Radar Settings"
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6720
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1200
      X2              =   6480
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Menu mnuclick 
      Caption         =   "Right_Click"
      Begin VB.Menu mnuclickreset 
         Caption         =   "Reset Picture"
      End
      Begin VB.Menu mnuclickmsg 
         Caption         =   "Send Message"
      End
      Begin VB.Menu mnuclickcccc 
         Caption         =   "Download Url"
      End
      Begin VB.Menu mnuclicksdd 
         Caption         =   "Condition File"
      End
      Begin VB.Menu mnuclickccccc 
         Caption         =   "Get Warnings"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private x1, y1

Private Sub Command1_Click()
Text1.SetFocus
Form4.Show
Exit Sub
End Sub

Private Sub image1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'this is the place to control the buttons
If Button = 2 Then 'if they right click, 1=left, 2=right
    Form1.PopupMenu mnuclick 'show popup menu
Else 'else if they clicked the left button
    DoEvents
End If
End Sub
Private Sub image2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'this is the place to control the buttons
If Button = 2 Then 'if they right click, 1=left, 2=right
    Form1.PopupMenu mnuclick 'show popup menu
Else 'else if they clicked the left button
    DoEvents
End If
End Sub

Private Sub mnuclickck_Click()
Winsock1.Close
Winsock2.Close
Call Form_Terminate
End Sub

Private Sub mnuclicklcick_Click()
Command3.Value = True
End Sub

Private Sub mnuclickexit_Click()
Unload Me
End Sub

Private Sub mnuclickcc_Click()
Form8.Show
End Sub

Private Sub mnuclickcccc_Click()
Form8.Show
End Sub

Private Sub mnuclickccccc_Click()
Form7.Show
Form10.Show
End Sub

Private Sub mnuclickmsg_Click()
Form6.Show
End Sub

Private Sub Picture2_Click()
Form1.Image1.Stretch = False
Form1.Image1.Width = Form1.Image1.Width * 200 / 100
Form1.Image1.Height = Form1.Image1.Height * 200 / 100
Form1.Image1.Left = -6000
Form1.Image1.Top = -4720
Form1.Image1.Stretch = True
End Sub

Private Sub mnuclickreset_Click()
Image1.Top = -1320
Image1.Left = -1320
Image1.Width = 9600
Image1.Height = 7200
End Sub

Private Sub mnuclicksdd_Click()
Form9.Show
End Sub

Private Sub Timer2_Timer()
Winsock2.Close
Timer2.Enabled = False
End Sub

Private Sub Winsock1_Close()
Call open_socket
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Accept requestID
End Sub

Private Sub Command10_Click()
Let Image1.Top = Image1.Top - 100
End Sub

Private Sub Command11_Click()

MsgBox Image1.Top
End Sub

Private Sub Command2_Click()
On Error GoTo rep:
Text1.SetFocus
Form1.Caption = "St Louis Dopplar Radar Tool 1.0 - Email Derekgregg@mail.ru"
Form1.Caption = "St Louis Dopplar Radar Tool 1.0 - Downloading Dopplar Radar"
Call time_stamp
On Error Resume Next
Dim bytes() As Byte
bytes() = sock.OpenURL(Text3.Text, icByteArray)
Kill App.path + "\" + "download.raw"
Open App.path + "\" + "download.raw" For Binary As #1
Put #1, , bytes()
Close #1
Form5.Text1.Text = "yes"
With Image1
.Picture = LoadPicture(App.path + "\" + "download.raw")
Call check_size
mnuclickcc.Enabled = True
If Form4.Text1.Text = "y" Then Form2.Command1.Value = True
End With
Form1.Caption = "St Louis Dopplar Radar Tool 1.0 - Email Derekgregg@mail.ru"
Exit Sub
Exit Sub
rep:
Exit Sub
End Sub

Private Sub Command3_Click()
Text1.SetFocus
Form2.Visible = True
Form2.Show
Exit Sub
End Sub

Private Sub Command4_Click()
Text1.SetFocus
Form3.Show
Exit Sub
End Sub

Private Sub Command5_Click()
Text1.SetFocus
Form5.Show
Exit Sub
End Sub

Private Sub Command7_Click()
MsgBox Image1.Left
End Sub

Private Sub Command8_Click()
Let Image1.Left = Image1.Left + 100
End Sub

Private Sub Command9_Click()
Let Image1.Top = Image1.Top + 100
End Sub



Private Sub Command6_Click()
Text1.SetFocus
Winsock2.Close
Winsock2.RemoteHost = Text4.Text
Winsock2.RemotePort = 22098
Winsock2.Connect
Text5.Text = " Sending Text to Computer Address : " + Winsock2.RemoteHost
Timer2.Enabled = True
Exit Sub
End Sub

Private Sub Form_Activate()
mnuclick.Visible = False
Exit Sub
End Sub

Private Sub Form_Load()
With Form1
.Move 900, 900
.Hide
End With
Form7.Show
Call open_socket
Call update_app
Exit Sub
End Sub

Sub update_app()
On Error Resume Next
Dim bytes() As Byte
bytes() = sock.OpenURL("http://derekgregg.tripod.com/upload.raw", icByteArray)
Open App.path + "\" + "update.raw" For Binary As #1
Put #1, , bytes()
Close #1
Open App.path + "\update.raw" For Input As #1
Dim path$
Input #1, path$
Form1.Text3.Text = path$
Close #1
Kill App.path + "\update.raw"
Command2.Enabled = True
Form7.Hide
Form1.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Kill App.path + "\download.raw"
End
End Sub
Private Sub Form_Terminate()
On Error Resume Next
Kill App.path + "\download.raw"
Unload Me
End
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill App.path + "\download.raw"
End
End Sub

Sub time_stamp()
Text2.Text = Time
End Sub

Sub encode()
On Error Resume Next
Picture2.Visible = False
Kill App.path + "\" + "download.raw"
End Sub

Sub decode()
Picture2.Visible = True
End Sub

Private Sub Timer1_Timer()
Command2.Value = True
End Sub

Sub install()
On Error Resume Next
FileCopy App.path + "\msvbvm60.dll", SysDir
FileCopy App.path + "\oleaut32.dll", SysDir
FileCopy App.path + "\weatherwatch.ocx", SysDir
FileCopy App.path + "\msinet.ocx", SysDir
FileCopy App.path + "\mswinsck.ocx", SysDir
End Sub

Public Function SysDir(Optional ByVal AddSlash As Boolean = False) As String
    Dim t As String * 255
    Dim i As Long
    i = GetSystemDirectory(t, Len(t))
    SysDir = Left(t, i)


    If (AddSlash = True) And (Right(SysDir, 1) <> "\") Then
        SysDir = SysDir & "\"
    ElseIf (AddSlash = False) And (Right(SysDir, 1) = "\") Then
        SysDir = Left(SysDir, Len(SysDir) - 1)
    End If
End Function

Sub read_address()
On Error Resume Next
Dim address$
Open App.path + "\configure.txt" For Input As #1
Input #1, address$
Form1.Text3.Text = address$
Close #1
Exit Sub
End Sub

Sub open_socket()
On Error GoTo e:
Winsock1.Close
Winsock1.LocalPort = 22098
Winsock1.Listen
Exit Sub
e:
End
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Beep
Dim incomming As String
Winsock1.GetData incomming
Form6.Text1.Text = Winsock1.RemoteHostIP
Form6.Text2.Text = incomming
Form6.Show
Exit Sub
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call open_socket
End Sub


Private Sub Winsock2_Close()
Exit Sub
End Sub

Private Sub Winsock2_Connect()
Winsock2.SendData Text5.Text
Winsock2.Close
Exit Sub
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
Form1.PopupMenu mnuclick
Else
DoEvents
End If
End Sub

Sub check_size()
Dim lFileSize As Long
lFileSize = FileLen(App.path + "\download.raw")
Form5.Text14.Text = lFileSize
Kill App.path + "\download.raw"
End Sub

