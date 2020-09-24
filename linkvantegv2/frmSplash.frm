VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3810
   ClientLeft      =   3810
   ClientTop       =   1710
   ClientWidth     =   5010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5010
   Begin VB.TextBox txtrealrefer 
      Height          =   285
      Left            =   0
      TabIndex        =   40
      Text            =   "http://linkvantage.com/Page2.asp?CampaignID="
      Top             =   6480
      Width           =   3615
   End
   Begin VB.TextBox txtreallink 
      Height          =   285
      Left            =   0
      TabIndex        =   39
      Text            =   "/link.asp?CampaignID="
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5160
      TabIndex        =   38
      Top             =   6480
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4080
      TabIndex        =   37
      Text            =   "&categoryid="
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox txtrefernum 
      Height          =   285
      Left            =   3600
      TabIndex        =   36
      Text            =   "0"
      Top             =   6480
      Width           =   495
   End
   Begin VB.TextBox txtlinknum 
      Height          =   285
      Left            =   2760
      TabIndex        =   35
      Text            =   "0"
      Top             =   6000
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   740
      ScaleHeight     =   3375
      ScaleWidth      =   4305
      TabIndex        =   29
      Top             =   360
      Visible         =   0   'False
      Width           =   4300
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3375
         Left            =   -1800
         Picture         =   "frmSplash.frx":030A
         ScaleHeight     =   3375
         ScaleWidth      =   6495
         TabIndex        =   30
         Top             =   0
         Width           =   6495
         Begin VB.Label Label9 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSplash.frx":932C
            Height          =   2175
            Left            =   1920
            TabIndex        =   31
            Top             =   120
            Width           =   4095
         End
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   740
      ScaleHeight     =   3495
      ScaleWidth      =   4305
      TabIndex        =   32
      Top             =   360
      Visible         =   0   'False
      Width           =   4300
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "http://www.linkvantage.com/signup.asp?referrer=XxSpyxX"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00AC7822&
         Height          =   255
         Left            =   720
         MouseIcon       =   "frmSplash.frx":940C
         MousePointer    =   99  'Custom
         TabIndex        =   34
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmSplash.frx":955E
         Height          =   1455
         Left            =   120
         TabIndex        =   33
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.TextBox winsockstatetxt 
      Height          =   285
      Left            =   1560
      TabIndex        =   28
      Top             =   3840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock6 
      Left            =   720
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox min1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   4560
      Picture         =   "frmSplash.frx":96F8
      ScaleHeight     =   195
      ScaleWidth      =   180
      TabIndex        =   27
      Top             =   15
      Width           =   187
   End
   Begin VB.PictureBox min2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   4560
      Picture         =   "frmSplash.frx":97CA
      ScaleHeight     =   195
      ScaleWidth      =   180
      TabIndex        =   26
      Top             =   15
      Visible         =   0   'False
      Width           =   187
   End
   Begin VB.PictureBox x2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   4800
      Picture         =   "frmSplash.frx":989B
      ScaleHeight     =   195
      ScaleWidth      =   180
      TabIndex        =   25
      Top             =   15
      Visible         =   0   'False
      Width           =   187
   End
   Begin VB.PictureBox x1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   4800
      Picture         =   "frmSplash.frx":9976
      ScaleHeight     =   195
      ScaleWidth      =   180
      TabIndex        =   24
      Top             =   15
      Width           =   187
   End
   Begin MSWinsockLib.Winsock Winsock5 
      Left            =   2280
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txttest 
      Height          =   285
      Left            =   240
      TabIndex        =   22
      Top             =   7320
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   3960
      TabIndex        =   21
      Top             =   6960
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.TextBox txtNewProxy 
      Height          =   285
      Left            =   5880
      TabIndex        =   20
      Top             =   6480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtProxy 
      Height          =   285
      Left            =   5880
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   840
      TabIndex        =   2
      Text            =   "0"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSWinsockLib.Winsock Winsock4 
      Left            =   1920
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   1560
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   1200
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtcookie 
      Height          =   285
      Left            =   0
      TabIndex        =   17
      Top             =   6960
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtrefersend 
      Height          =   285
      Left            =   0
      TabIndex        =   15
      Top             =   5520
      Width           =   6135
   End
   Begin VB.TextBox txtlinksend 
      Height          =   285
      Left            =   0
      TabIndex        =   14
      Top             =   5160
      Width           =   6135
   End
   Begin VB.TextBox txtRecievedData 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.TextBox usertxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Enter a user name"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox fnametxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Enter a password"
      Top             =   840
      Width           =   1095
   End
   Begin VB.Timer timeSend 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   840
      Top             =   1560
   End
   Begin VB.TextBox Texttime 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Text            =   "1"
      ToolTipText     =   "Enter an interval in seconds"
      Top             =   1440
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   840
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   0
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   4500
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      Caption         =   "  LinkVantage v2.0   b y   X x S p y x X"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      MouseIcon       =   "frmSplash.frx":9A51
      MousePointer    =   3  'I-Beam
      TabIndex        =   23
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      MouseIcon       =   "frmSplash.frx":9BA3
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label lblstatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00AC7822&
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   16
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmSplash.frx":9CF5
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmSplash.frx":9E47
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   480
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   720
      X2              =   720
      Y1              =   360
      Y2              =   3720
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      MouseIcon       =   "frmSplash.frx":9F99
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1200
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2880
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label lblstat 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Interval"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clicks:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Command2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stop"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      MouseIcon       =   "frmSplash.frx":A0EB
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For dragging the form
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const SW_SHOWMAXIMIZED = 3

Private bLeftOut As Boolean
Private bRightOut As Boolean
Private bBottomOut As Boolean

Private iRelLeftTrayOffset As Integer
Private iRelRightTrayOffset As Integer
Private iRelBottomTrayOffset As Integer

Private Declare Sub ReleaseCapture Lib "User32" ()
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "User32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal x3 As Integer, ByVal y3 As Integer) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Private Sub Form_Load()
'On Error GoTo hell:
 
 
 
 lblstatus.Caption = "Status: Ready...."
 
 
'hell:
'MsgBox ("Missing files")
'Unload Me
 

End Sub

Private Sub Label11_Click()
On Error Resume Next
gotoweb
End Sub

Private Sub Label3_Click()
Winsock6.Connect "www.linkvantage.com", "80"
End Sub

Private Sub Label6_Click()

If Picture1.Visible = True Then
Picture1.Visible = False
Picture3.Visible = False
Else
Picture1.Visible = True
Picture2.Visible = True
Picture3.Visible = False
End If

If Picture3.Visible = True Then
Picture3.Visible = False
Picture1.Visible = True
Picture2.Visible = True
End If
End Sub

Private Sub Label7_Click()
If Picture3.Visible = True Then
Picture3.Visible = False
Picture1.Visible = False
Else
Picture3.Visible = True
Picture1.Visible = False
End If
End Sub







'Close Code
Private Sub x2_Click()
Unload Me
End Sub

Private Sub x1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    x1.Visible = False
    x2.Visible = True

End Sub
Private Sub x1_Click()
    x2_Click
End Sub
Private Sub x2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub
'end close code

'Min code

Private Sub min2_Click()
Call AddToTray(Me.Icon, Me.Caption, Me)
End Sub

Private Sub min1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    min1.Visible = False
    min2.Visible = True

End Sub
Private Sub min1_Click()
    min2_Click
End Sub
Private Sub min2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call AddToTray(Me.Icon, Me.Caption, Me)
End Sub
'end min code



Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' move the window
    ReleaseCapture
    SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub


Private Sub Winsock1_Connect()
Dim strrefer As String
Dim strcookie As String

lblstatus.Caption = "Status: Clicking links"

strcookie = txtNewProxy.Text & "; LinkVantage=Password=" & fnametxt & "&Username=" & usertxt & ""
strrefer = txtrefersend.Text


strData = "GET " & txtlinksend.Text & " HTTP/1.1" & vbCrLf
strData = strData & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-gsarcade-launch, */*" & vbCrLf
strData = strData & "Referer: " + strrefer + vbCrLf
strData = strData & AL() & vbCrLf
strData = strData & "Accept-Encoding: gzip, deflate" & vbCrLf
strData = strData & UA() & vbCrLf
strData = strData & "Host: linkvantage.com" & vbCrLf
strData = strData & "Connection: Keep-Alive" & vbCrLf
strData = strData & "Cookie: " + strcookie + vbCrLf & vbCrLf


Winsock1.SendData strData
End Sub


Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim RecievedData As String
Winsock1.GetData RecievedData
txtRecievedData.Text = txtRecievedData.Text & RecievedData
    
    ' Get Cookie
    If txtcookie.Text = "" Then txtcookie.Text = GetCookie(txtRecievedData.Text)
    
    txtRecievedData.Text = ""
    
End Sub


Private Sub Winsock1_SendComplete()


Text1.Text = Text1.Text + 1
lblstat.Caption = lblstat.Caption + 1

txtRecievedData.Text = ""

End Sub

Private Sub Winsock6_Connect()
Dim strData As String
lblstatus.Caption = "Status: Getting Cookie"

strData = "GET / HTTP/1.1" & vbCrLf
strData = strData & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-gsarcade-launch, */*" & vbCrLf
strData = strData & AL() & vbCrLf
strData = strData & "Accept-Encoding: gzip, deflate" & vbCrLf
strData = strData & UA() & vbCrLf
strData = strData & "Host: www.linkvantage.com" & vbCrLf
strData = strData & "Connection: Keep-Alive" & vbCrLf


Winsock6.SendData strData & vbCrLf


End Sub


Private Sub Winsock6_DataArrival(ByVal bytesTotal As Long)
Dim RecievedData As String
Winsock6.GetData RecievedData
txtRecievedData.Text = txtRecievedData.Text & RecievedData
    
    ' Get Cookie
    If txtcookie.Text = "" Then
    txtcookie.Text = GetCookie(txtRecievedData.Text)
    End If
    
winsockstatetxt.Text = Winsock2.State

     If winsockstatetxt.Text = "7" Then
     MsgBox ("already connected")
   Exit Sub
End If

Winsock2.Connect "www.linkvantage.com", "80"
  
End Sub

Private Sub Winsock2_Connect()
Dim strData As String
Dim LengthBody As Integer
Dim Body As String
Dim strrefer As String
Dim strcookie As String

lblstatus.Caption = "Status: Sending login information"

strcookie = txtcookie.Text
strrefer = "http://linkvantage.com/Main.asp"

LengthBody = Len(Body)

strData = "POST /SetCookie.asp?Username=" & usertxt.Text & "&password=" & fnametxt.Text & " HTTP/1.1" & vbCrLf
strData = strData & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-gsarcade-launch, */*" & vbCrLf
strData = strData & AL() & vbCrLf
strData = strData & "Referer: " + strrefer + vbCrLf
strData = strData & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
strData = strData & "Accept-Encoding: gzip, deflate" & vbCrLf
strData = strData & UA() & vbCrLf
strData = strData & "Host: linkvantage.com" & vbCrLf
strData = strData & "Content-Length: " & LengthBody & vbCrLf
strData = strData & "Connection: Keep-Alive" & vbCrLf
strData = strData & "Cache-Control: no-cache" & vbCrLf
strData = strData & "Cookie: " + strcookie + vbCrLf & vbCrLf

Winsock2.SendData strData


End Sub

Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
Dim RecievedData As String
Winsock2.GetData RecievedData
txtRecievedData.Text = txtRecievedData.Text & RecievedData
    
   ' Get Cookie
    txtcookie.Text = GetCookie(txtRecievedData.Text)
    
    If txtcookie.Text = "0 Continue" Then txtcookie.Text = GetCookie(txtRecievedData.Text)
    
     
winsockstatetxt.Text = Winsock3.State

     If winsockstatetxt.Text = "7" Then
   Exit Sub
End If

Winsock3.Connect "www.linkvantage.com", "80"


End Sub


Private Sub Winsock2_SendComplete()
'txtRecievedData.Text = ""
End Sub

Private Sub Winsock3_Connect()


GoTo Parse:

Parse:

txttest.Text = txtcookie.Text

Dim proxy As String
proxy = txttest.Text
txtProxy.Text = txttest.Text
Dim first As Variant
Dim second As Variant
first = InStr(txtProxy.Text, ";")
txtNewProxy.Text = Left(txtProxy.Text, (first - 1))
second = Len(txtProxy.Text)
txtPort.Text = Right(txtProxy.Text, (second - first))
If txttest.Text = "" Then
GoTo Parse:
Else

lblstatus.Caption = "Status: Confirming Cookie"

Dim strData As String
Dim LengthBody As Integer
Dim Body As String
Dim strrefer As String
Dim strcookie As String
txtRecievedData.Text = ""

strcookie = txtNewProxy.Text & "; LinkVantage=Password=" & fnametxt & "&Username=" & usertxt & ""
strrefer = "http://linkvantage.com/Main.asp"

LengthBody = Len(Body)

strData = "POST /SetCookie.asp?try=Cookie&Username=" & usertxt.Text & "&password=" & fnametxt.Text & " HTTP/1.1" & vbCrLf
strData = strData & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-gsarcade-launch, */*" & vbCrLf
strData = strData & AL() & vbCrLf
strData = strData & "Referer: " + strrefer + vbCrLf
strData = strData & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
strData = strData & "Accept-Encoding: gzip, deflate" & vbCrLf
strData = strData & UA() & vbCrLf
strData = strData & "Host: linkvantage.com" & vbCrLf
strData = strData & "Content-Length: " & LengthBody & vbCrLf
strData = strData & "Connection: Keep-Alive" & vbCrLf
strData = strData & "Cookie: " + strcookie + vbCrLf & vbCrLf
strData = strData & "Cache-Control: no-cache" & vbCrLf

Winsock3.SendData strData
End If
End Sub

Private Sub Winsock3_DataArrival(ByVal bytesTotal As Long)
Dim RecievedData As String
Winsock3.GetData RecievedData
txtRecievedData.Text = txtRecievedData.Text & RecievedData
    
   ' Get Cookie
    'txtcookie.Text = GetCookie(txtRecievedData.Text)
    'If txtcookie.Text = "0 Continue" Then txtcookie.Text = GetCookie(txtRecievedData.Text)
    ' txtRecievedData.Text = ""
    
     winsockstatetxt.Text = Winsock4.State

     If winsockstatetxt.Text = "7" Then
   Exit Sub
End If

Winsock4.Connect "www.linkvantage.com", "80"
     
   txtRecievedData.Text = ""

End Sub


Private Sub Winsock4_Connect()
Dim strData As String
Dim LengthBody As Integer
Dim Body As String
Dim strrefer As String
Dim strcookie As String

lblstatus.Caption = "Status: Processing logon"

strcookie = txtNewProxy.Text & "; LinkVantage=Password=" & fnametxt & "&Username=" & usertxt & ""
strrefer = "http://linkvantage.com/SuccessLogin.asp"

LengthBody = Len(Body)

strData = "POST /ProcessLogin.asp HTTP/1.1" & vbCrLf
strData = strData & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-gsarcade-launch, */*" & vbCrLf
strData = strData & AL() & vbCrLf
strData = strData & "Referer: " + strrefer + vbCrLf
strData = strData & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
strData = strData & "Accept-Encoding: gzip, deflate" & vbCrLf
strData = strData & UA() & vbCrLf
strData = strData & "Host: linkvantage.com" & vbCrLf
strData = strData & "Content-Length: " & LengthBody & vbCrLf
strData = strData & "Connection: Keep-Alive" & vbCrLf
strData = strData & "Cookie: " + strcookie + vbCrLf & vbCrLf


Winsock4.SendData strData

End Sub

Private Sub Winsock4_DataArrival(ByVal bytesTotal As Long)

Dim RecievedData As String
Winsock4.GetData RecievedData
txtRecievedData.Text = txtRecievedData.Text & RecievedData
    
   ' Get Cookie
    If txtcookie.Text = "" Then txtcookie.Text = GetCookie(txtRecievedData.Text)
    If txtcookie.Text = "0 Continue" Then txtcookie.Text = GetCookie(txtRecievedData.Text)
    
   winsockstatetxt.Text = Winsock5.State

If InStr(txtRecievedData.Text, "LoginFail") <> 0 Then
 lblstatus.Caption = "Status: Pass or Login Invalid"
 Winsock6.Close
 Winsock2.Close
 Winsock3.Close
 Winsock4.Close
 Winsock5.Close
 Command1.Visible = False
Command2.Visible = False
Label3.Visible = True
 Exit Sub
 Else
     If winsockstatetxt.Text = "7" Then
   Exit Sub
     End If
End If

Winsock5.Connect "www.linkvantage.com", "80"
        

   
End Sub
     
     Private Sub Winsock5_Connect()
Dim strData As String
Dim LengthBody As Integer
Dim Body As String
Dim strrefer As String
Dim strcookie As String
 
If InStr(txtRecievedData.Text, "LoginFail") <> 0 Then
 lblstatus.Caption = "Status: Pass or Login Invalid"
 Winsock6.Close
 Winsock2.Close
 Winsock3.Close
 Winsock4.Close
 Winsock5.Close
 Command1.Visible = False
Command2.Visible = False
Label3.Visible = True
 Exit Sub
 Else

lblstatus.Caption = "Status: Information Valid"

strcookie = txtNewProxy.Text & "; LinkVantage=Password=" & fnametxt & "&Username=" & usertxt & ""
strrefer = "http://linkvantage.com/Main.asp"

LengthBody = Len(Body)

strData = "POST /SuccessLogin.asp HTTP/1.1" & vbCrLf
strData = strData & "Accept: image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-gsarcade-launch, */*" & vbCrLf
strData = strData & AL() & vbCrLf
strData = strData & "Referer: " + strrefer + vbCrLf
strData = strData & "Content-Type: application/x-www-form-urlencoded" & vbCrLf
strData = strData & "Accept-Encoding: gzip, deflate" & vbCrLf
strData = strData & UA() & vbCrLf
strData = strData & "Host: linkvantage.com" & vbCrLf
strData = strData & "Content-Length: " & LengthBody & vbCrLf
strData = strData & "Connection: Keep-Alive" & vbCrLf
strData = strData & "Cookie: " + strcookie + vbCrLf & vbCrLf


Winsock5.SendData strData
End If
End Sub

Private Sub Winsock5_SendComplete()
Label3.Visible = False
Command1.Visible = True

txtRecievedData.Text = ""
End Sub

Private Sub Winsock5_DataArrival(ByVal bytesTotal As Long)
If winsockstatetxt.Text = "0" Then
   Exit Sub
     End If
Dim RecievedData As String
Winsock5.GetData RecievedData
txtRecievedData.Text = txtRecievedData.Text & RecievedData


 lblstatus.Caption = "Status: Logged in"
txtRecievedData.Text = ""
 

End Sub


Private Sub Command1_Click()
If Texttime.Text = "Interval" Then
MsgBox ("please enter a interval 1-10")
Exit Sub
End If
If usertxt.Text = "" Then
MsgBox ("please enter a valid username")
Exit Sub
End If
If fnametxt.Text = "" Then
MsgBox ("please enter your Password")
Exit Sub
End If
timeSend.Interval = (Texttime.Text & "000")
timeSend.Enabled = True
lblstatus.Caption = "Status: Starting"

Command1.Visible = False
Command2.Visible = True
End Sub

Private Sub Command2_Click()
timeSend.Enabled = False
Command1.Visible = False
Command2.Visible = False
Label3.Visible = True
Winsock6.Close
 Winsock2.Close
 Winsock3.Close
 Winsock4.Close
 Winsock1.Close
 Winsock5.Close
 lblstatus.Caption = "Status: Stopped"
End Sub

Private Sub timeSend_Timer()

On Error Resume Next

'linklist.ListIndex = linklist.ListIndex + 1
'referlist.ListIndex = referlist.ListIndex + 1
Text3.Text = Int(1000 * Rnd)
txtlinknum.Text = txtlinknum.Text + 1
txtrefernum.Text = txtrefernum.Text + 1
txtlinksend = txtreallink.Text + txtlinknum.Text
txtrefersend = txtrealrefer.Text + txtrefernum.Text + Text2.Text + Text3.Text
Winsock1.Close
Winsock1.Connect "www.linkvantage.com", "80"



End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Command1.FontUnderline = True
Command1.ForeColor = vbBlack
Command1.BackColor = vbWhite

End Sub
Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Command2.FontUnderline = True
Command2.ForeColor = vbBlack
Command2.BackColor = vbWhite

End Sub
Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label3.FontUnderline = True
Label3.ForeColor = vbBlack
Label3.BackColor = vbWhite

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Command1.FontUnderline = False
Command1.BackColor = &H404040
Command1.ForeColor = vbWhite
Command2.FontUnderline = False
Command2.BackColor = &H404040
Command2.ForeColor = vbWhite
Label3.FontUnderline = False
Label3.BackColor = &H404040
Label3.ForeColor = vbWhite
   If RespondToTray(X) <> 0 Then Call ShowFormAgain(Me)
End Sub

Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.FontUnderline = False
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Label11.FontUnderline = True


End Sub
