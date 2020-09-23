VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form frmhotmail 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hotmail Messages"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   Icon            =   "frmhotmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmhotmail.frx":08CA
   ScaleHeight     =   5220
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin SocketWrenchCtrl.Socket Socket 
      Left            =   1800
      Top             =   4800
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   -1  'True
      Backlog         =   5
      Binary          =   -1  'True
      Blocking        =   -1  'True
      Broadcast       =   0   'False
      BufferSize      =   0
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   0
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.CommandButton cmdFinish 
      Caption         =   "Finish"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   32
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CheckBox chkSendSig 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Always Send Signature with Outgoing Mail"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   31
      Top             =   3240
      Width           =   3855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   30
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "Modify"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   29
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox cboLogin 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      TabIndex        =   27
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddAccount 
      Caption         =   "Add Account"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   26
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtSignature 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   240
      MaxLength       =   500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   3840
      Width           =   3735
   End
   Begin VB.CheckBox chkSaveSent 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Always Save Outgoing Messages"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   3600
      Width           =   3855
   End
   Begin VB.CheckBox chkNoSounds 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Don't Play Sounds"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   2880
      Width           =   3735
   End
   Begin VB.CheckBox chkNoDlgFocus 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Don't Focus on Dialogs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   2520
      Width           =   3735
   End
   Begin VB.TextBox txtTimeout 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6480
      TabIndex        =   18
      Text            =   "3"
      Top             =   2160
      Width           =   615
   End
   Begin VB.CheckBox chkTimeout 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Connection Timeout"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   17
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Timer timeout 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3720
      Top             =   4800
   End
   Begin VB.Timer nettimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   3240
      Top             =   4800
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   6
      Left            =   1200
      Picture         =   "frmhotmail.frx":80BD6
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   5
      Left            =   3600
      Picture         =   "frmhotmail.frx":818A0
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.CheckBox chkStartup 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Run Hotmail Checker at Windows Startup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   1800
      Width           =   3735
   End
   Begin VB.CheckBox chkAutoLogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Automatic Login"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   1440
      Width           =   3615
   End
   Begin VB.CheckBox chkRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check Inbox Every"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox txtInterval 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   10
      Text            =   "1"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CheckBox chkAllDlgs 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Show Notification Dialogs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   720
      Width           =   3615
   End
   Begin VB.Timer blinker 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2280
      Top             =   4800
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   4
      Left            =   3000
      Picture         =   "frmhotmail.frx":8216A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   2400
      Picture         =   "frmhotmail.frx":82A34
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer refresher 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2760
      Top             =   4800
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   1800
      Picture         =   "frmhotmail.frx":832FE
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "frmhotmail.frx":83740
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIcon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   600
      Picture         =   "frmhotmail.frx":83B82
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picNotify 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      Picture         =   "frmhotmail.frx":8444C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtpass 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   240
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2880
      Width           =   2220
   End
   Begin VB.Label lblsignin 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sign-In Name: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   240
      TabIndex        =   28
      Top             =   1920
      Width           =   1410
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Signature:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   4320
      X2              =   8400
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblFolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4320
      TabIndex        =   23
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "minutes."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   19
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "minutes."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   3975
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
      BeginProperty Font 
         Name            =   "News Gothic MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   4200
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   3930
   End
   Begin VB.Label lblpassword 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Menu mnuTray 
      Caption         =   "Tray"
      Visible         =   0   'False
      Begin VB.Menu itmHotmail 
         Caption         =   "&Hotmail.com"
      End
      Begin VB.Menu itmCheck 
         Caption         =   "Check Mail"
      End
      Begin VB.Menu itmSep1 
         Caption         =   "-"
      End
      Begin VB.Menu itmShowSettings 
         Caption         =   "&Settings"
      End
      Begin VB.Menu itmClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "frmhotmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iconnum As Integer
Dim minutes As Integer
Dim closedbyTray As Boolean
Dim gotFolders As Boolean
Dim pageNum As Integer

Private Sub blinker_Timer()
    If iconnum = 1 Then
        iconnum = 2
        ChangeIcon picIcon(3).Picture, picNotify
    ElseIf iconnum = 2 Then
        iconnum = 1
        ChangeIcon picIcon(4).Picture, picNotify
    End If
End Sub

Private Sub cboLogin_Click()
    Dim choice As Integer
    
    txtpass = RTrim(Accounts(cboLogin.ListIndex).password)
    Socket.Action = SOCKET_DISCONNECT
    GotMail = False
    ResetAll
    txtpass.SetFocus
    ConnectToHotmail
End Sub

Private Sub chkAllDlgs_Click()
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "ShowDlgs", chkAllDlgs.Value
    If chkAllDlgs.Value = 1 Then
        ShowDlgs = True
    Else
        ShowDlgs = False
    End If
End Sub

Private Sub chkAutoLogin_Click()
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "SaveLogin", chkAutoLogin.Value
End Sub

Private Sub chkNoDlgFocus_Click()
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "NoDialogFocus", chkNoDlgFocus.Value
End Sub

Private Sub chkNoSounds_Click()
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "NoSounds", chkNoSounds.Value
End Sub

Private Sub chkRefresh_Click()
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Refresh", chkRefresh.Value
    If chkRefresh.Value = 1 Then
        REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Interval", txtInterval
    End If
End Sub

Private Sub chkSaveSent_Click()
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "SaveSent", chkSaveSent.Value
End Sub

Private Sub chkSendSig_Click()
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "SendSignature", chkSendSig.Value
End Sub

Private Sub chkStartup_Click()
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Startup", chkStartup.Value
    If chkStartup.Value = 1 Then
        REGSaveSetting vHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Hotmail Checker", App.Path & "\hmchecker.exe"
    Else
        DeleteValue vHKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "Hotmail Checker"
    End If
End Sub

Private Sub chkTimeout_Click()
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Timeout", chkTimeout.Value
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "TimeoutValue", txtTimeout
End Sub

Private Sub cmdAddAccount_Click()
    AddAccount cboLogin, cboLogin, txtpass
    SaveAccounts
End Sub

Private Sub cmdDelete_Click()
    Accounts(cboLogin.ListIndex).username = ""
    SaveAccounts
    LoadAccounts
    RefreshAccounts
End Sub

Private Sub cmdFinish_Click()
    Dim fnum As Integer
    fnum = FreeFile
    
    ' Save the signature file
    If Dir(App.Path & "\signature.txt") <> "" Then Kill (App.Path & "\signature.txt")
    Open App.Path & "\signature.txt" For Binary Access Write As #fnum 'Len = Len(MailData)
    Put #fnum, , txtSignature.text
    Close #fnum
    
    SaveAccounts
    
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Username", cboLogin.text
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Password", txtpass
            
    If chkRefresh.Value = 1 Then
        refresher = True
    End If
    Hide
End Sub

Private Sub cmdModify_Click()
    With Accounts(cboLogin.ListIndex)
        .loginname = cboLogin
        .password = txtpass
        .username = cboLogin
    End With
End Sub

Private Sub Form_Load()
    Dim temp As String * 500
    Dim fnum As Integer
    
    'Initialize Socket
    Socket.AddressFamily = AF_INET
    Socket.Binary = False
    Socket.Blocking = False
    Socket.BufferSize = 10000
    Socket.Protocol = IPPROTO_IP
    Socket.SocketType = SOCK_STREAM
    Socket.RemotePort = 80
    
    ' Load the signature file
    fnum = FreeFile
    If Dir(App.Path & "\signature.txt") <> "" Then
        Open App.Path & "\signature.txt" For Binary Access Read As #fnum 'Len = Len(MailData)
        Get #fnum, , temp
        txtSignature = RTrim(temp)
        Close #fnum
    End If
    
    ' Load accounts
    LoadAccounts
    RefreshAccounts
    
    ' Read in registry settings
    txtInterval = REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Interval")
    If txtInterval = "" Then txtInterval = "2"
    txtTimeout = REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "TimeoutValue")
    If txtTimeout = "" Then txtTimeout = "2"
    
    On Error Resume Next
    chkAllDlgs.Value = CInt(REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "ShowDlgs"))
    chkRefresh.Value = CInt(REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Refresh"))
    chkAutoLogin.Value = CInt(REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "SaveLogin"))
    chkStartup.Value = CInt(REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Startup"))
    chkNoSounds.Value = CInt(REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "NoSounds"))
    chkTimeout.Value = CInt(REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Timeout"))
    chkNoDlgFocus.Value = CInt(REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "NoDialogFocus"))
    chkSendSig.Value = CInt(REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "SendSignature"))
    chkSaveSent.Value = CInt(REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "SaveSent"))
    
    ' Load stored messages
    frmMailbox.LoadMessages
        
    If REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Configured") <> "1" Then
        lblInfo.Visible = True
    Else
        If REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "SaveLogin") = "1" Then
            ' Retrieve Login Info
            StrLogin = REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Username")
            StrPass = REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Password")
            If StrLogin <> "" And StrPass <> "" Then
                cboLogin.text = StrLogin
                txtpass = StrPass
                ' Hide the client and create the tray icon
                Hide
                CreateIcon picNotify, "Hotmail Checker"
                If IsNetConnectOnline = False Then
                    ShowTip "Internet Not Available.  Waiting for Connection..."
                    ChangeTip "Hotmail Checker NetDetect"
                    ChangeIcon picIcon(6).Picture, picNotify
                    nettimer = True
                End If
                
                ' Start connecting
                ConnectToHotmail
            Else
                frmhotmail.Show
            End If
        Else
            frmhotmail.Show
        End If
    End If
    
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Configured", "1"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim fnum As Integer
    fnum = FreeFile
    
    ' Save the signature file
    If Dir(App.Path & "\signature.txt") <> "" Then Kill (App.Path & "\signature.txt")
    Open App.Path & "\signature.txt" For Binary Access Write As #fnum 'Len = Len(MailData)
    Put #fnum, , txtSignature.text
    Close #fnum
    
    SaveAccounts
    
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Username", cboLogin.text
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Password", txtpass
    
    If Not closedbyTray Then
        Cancel = 1
        If chkRefresh.Value = 1 Then
            refresher = True
        End If
        Hide
    End If
End Sub

Private Sub Form_Resize()
    'If WindowState = vbMinimized Then Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Save stored messages
    frmMailbox.SaveMessages
    
    Unload frmMailbox
    Unload frmTip
    Unload frmNotify
    Unload frmFolders
    
    DeleteIcon picNotify
End Sub

Private Sub itmCheck_Click()
    'If GotMail Then
        CheckMail
    'End If
End Sub

Private Sub itmClose_Click()
    closedbyTray = True
    Unload frmTip
    Unload Me
End Sub

Private Sub itmHotmail_Click()
    ShellExecute hwnd, "Open", "http://www.hotmail.com", "", App.Path, 1
End Sub

Private Sub itmShowSettings_Click()
    'refresher = False
    'GotMail = False
    Show
    pSetForegroundWindow hwnd
End Sub

Private Sub lbl_Change()
    'ShowTip lbl
    ChangeTip lbl
End Sub

Private Sub nettimer_Timer()
    If IsNetConnectOnline Then
        nettimer = False
        ChangeTip "Hotmail Checker"
        ConnectToHotmail
    End If
End Sub

Private Sub picNotify_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    X = X / Screen.TwipsPerPixelX
    Select Case X
    Case WM_LBUTTONDOWN
        
    Case WM_RBUTTONUP
        PopupMenu mnuTray
    Case WM_MOUSEMOVE
        
    Case WM_LBUTTONDBLCLK
        'refresher = False
        'GotMail = False
        frmMailbox.Show
        pSetForegroundWindow frmMailbox.hwnd
        'frmMailbox.RefreshMessages
    End Select
End Sub

Private Sub refresher_Timer()
    If chkRefresh.Value = 1 Then
        'PlayWave App.Path & "\tic.wav"
        'ShowTip "TICK!"
        minutes = minutes + 1
        If CStr(minutes) = txtInterval Then
            If GotMail = True Then
                CheckMail
            End If
            minutes = 0
        End If
    End If
End Sub
' This sub will request the inbox without logging in again
Public Sub CheckMail()
    ResetIsOnline
    Socket.Action = SOCKET_DISCONNECT
    NextPage = NewUrl
    newmessages = ""
    loggedin = False
    ' Refresh the mailbox list
    frmMailbox.RefreshMessages
   
    ' Disable the new mail timer
    refresher = False
    GotMail = False
    timeout = True
    ChangeIcon picIcon(0).Picture, picNotify
    ChangeTip "Checking for New Mail..."
    ShowTip "Checking for New Mail..."
    iconnum = 1
    blinker = True
    BatchNumber = 1
    Socket.Action = 2
End Sub
' Resets all messages so they are not displayed if they are no longer present
Private Sub ResetIsOnline()
    Dim i As Integer
    For i = 0 To MsgCount - 1
        Messages(i).isonline = False
    Next
End Sub

Private Sub SOCKET_CONNECT()
    Dim str As String ' holds data to be sent to server
    GotMail = False
    
    Select Case BatchNumber
        Case 0
            blinker = True
            iconnum = 1
            
            lbl.Caption = "2. Sending Login Data..."
            ShowTip "Checking for New Mail..."
            str$ = MakeString(0) ' make first batch of data to send
        Case 1
            ' Reset the and message array
            'ReDim Messages(0) As HotmailMsg
            'MsgCount = 0
            
            lbl.Caption = "4. Requesting Mailbox..."
            str$ = MakeString(1) ' make second batch of data
        Case 2
            ' Get Inbox
            str = MakeString(2)
        Case 3
            ' Get Next Page
            str = MakeString(3, NextPage)
            BatchNumber = 2
        Case 4
            ' Get Folders
            lbl = "Fetching Folder List..."
            str = MakeString(4)
        Case 5
            GotMail = True
            ' Get Message
            lbl = "Fetching Message..."
            str = MakeString(5, msgURL)
        Case 6
            ' Send Message
            lbl = "Sending Message..."
            str = composeString
        Case 7
            ' Get compose page
            ' We have to do this before we can send
            ' because this page contains vital information
            lbl = "Fetching Compose Page..."
            ' We'll use the GetNextPage string.  I've added
            ' the Referer: header to it.
            str = MakeString(3, composeurl)
        Case 8
            lbl = "Deleting Message..."
            str = MakeString(6, msgURL)
    End Select
    
    ' send data to server
    Socket.SendLen = Len(str$)
    Socket.SendData = str$
End Sub

Private Sub Socket_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    MsgBox ErrorString & vbCrLf & vbCrLf & "HotmailBox will Re-Connect", vbCritical, "Socket Error"
    Socket.Action = SOCKET_DISCONNECT
    ConnectToHotmail
End Sub

Private Sub Socket_Read(DataLength As Integer, IsUrgent As Integer)
Dim fnum As Integer
Dim NewData As String ' holds the data we receive from hotmail server

Socket.RecvLen = DataLength
NewData = Socket.RecvData ' get data
Debug.Print NewData

Select Case BatchNumber ' depending on which batch of data we receive, we will do different actions
Case 0
    
    If InStr(1, NewData, "Location:") <> 0 Then ' in first batch, if login and password is correct, server directs you to a new server and new url
        Dim temp As String
        temp$ = Right(NewData, Len(NewData) - Len("Location: "))
        temp$ = Left(temp, Len(temp) - 2)
        NewHost = Mid(temp, 8, (Len(temp) - 8) - (Len(temp) - InStr(8, temp, "/"))) ' here we get the new server address
        NewUrl = Right(temp, Len(temp) - InStr(8, temp, "/")) ' and here we get the new url to request
        BatchNumber = 1
        lbl.Caption = "3. Finding Mailbox Server..."
        ' disconnect and reconnect to new server to send data
        Socket.Action = SOCKET_DISCONNECT
        Socket.HostName = NewHost
        Socket.Action = 2 ' once we connect, we'll request the new page (NewUrl)
    End If
    If InStr(1, NewData, "reauthhead.asp") <> 0 Then
        Socket.Action = SOCKET_DISCONNECT
        Show
        ShowNotify "Invalid Login or Password", picIcon(5).Picture

        blinker = False
        iconnum = 0
        ChangeIcon picIcon(5).Picture, picNotify
        lbl.Caption = "Error: Invalid Login or Password"
        Call ResetAll
    End If
Case 1
    If InStr(1, NewData, "Set-Cookie:") <> 0 Then ' now that we've succesfully sent the correct data to the new server, it sends cookies to be re-sent when we request the mailbox
        Dim c1 As Integer, c2 As Integer
        c1 = InStr(1, NewData, "Set-Cookie: ")
        If c1 <> 0 Then c1 = c1 + Len("Set-Cookie: ")
        c2 = InStr(1, NewData, ";")
        If c2 = 0 Then c2 = Len(NewData)
        Cookies(CurrentCookie) = Mid(NewData, c1, c2 - c1)
        'Debug.Print Cookies(CurrentCookie)
        CurrentCookie = CurrentCookie + 1
    End If
    If InStr(1, NewData, "Refresh") <> 0 Then 'after the server sends all the cookies, it tell us to refresh to the actual mailbox, therefore demanding us to send the cookies back to the server
        NewUrl = Mid(NewData, InStr(1, NewData, "content=") + 16, Len(NewData) - (InStr(1, NewData, "content=") + 16) - 3) ' the url of the final mailbox
        c1 = InStr(1, NewUrl, "/cgi")
        If c1 <> 0 Then
            NewUrl = Mid(NewUrl, c1, (Len(NewUrl) - c1) + 1)
        End If
        'BatchNumber = 2
        'Dim str As String
        'str$ = MakeString(2) ' compile the final data to be sent, containing the url of the mailbox, and all of the cookies received
        ' now send the data
        'Socket.SendLen = Len(str$)
        'Socket.SendData = str$ ' send final data
    End If
    If InStr(1, NewData, "</html>") <> 0 Then
        BatchNumber = 2
        Socket.Action = SOCKET_DISCONNECT
        Socket.Action = 2
    End If
    If InStr(1, NewData, "reason=nocookies") <> 0 Then
        ShowTip "Could not connect.  Retrying..."
        Socket.Action = SOCKET_DISCONNECT
        ResetAll
        ConnectToHotmail
    End If
Case 2 ' if all correct data was send correctly, on the third time we begin to receive the mailbox data
    'lbl.Caption = "5. Processing Mailbox..."
    If InStr(1, NewData, "<html>") <> 0 Then ' here is where the number of new messages can be read
        ReadBox = True ' begin storing incoming batches (pages) into the variable 'BoxData'
        BoxBatch = 0 ' we will only store 10 batches of data, as that is all we need to find the new messages
        MailData = NewData
    End If
    If ReadBox = True Then
        BoxBatch = BoxBatch + 1
        MailData = MailData & NewData
        If InStr(1, NewData, "</html>") <> 0 Then GoTo 1
    End If
    Exit Sub
1: ' we now have all the crucial mailbox source stored, and we are ready to extract the number of new messages from it.
   ' By storing more batches, you can also extract other information that you want. this is just shown as an example.
    Socket.Action = SOCKET_DISCONNECT
    Dim Location As Integer, Offset As Integer, Length As Integer
    
    ' Extract all messages from the current page
    temp = FindNextMessage(lastmsg + 1, MailData)
    Do Until temp = ""
        'lstMsgs.AddItem temp
        temp = FindNextMessage(lastmsg + 1, MailData)
    Loop
    lastmsg = 0
        
    frmMailbox.RefreshMessages
    
    If NextPage = "" Then NextPage = NewUrl
    CurrentPage = NextPage
    
    If newmessages = "" Then
        Location = InStr(1, MailData, "new")
        Offset = InStr(Location - 5, MailData, ">") + 1
        Length = Location - (Location - Offset)
        newmessages = Mid(MailData, Length, Len(MailData) - Offset - (Len(MailData) - Location) - 1) ' store value of new messages
        If Int(newmessages) > 1 Then
            newmessages = "You have " & newmessages & " new messages"
        ElseIf Int(newmessages) = 1 Then
            newmessages = "You have " & newmessages & " new message"
        Else
            newmessages = "You have no new messages"
        End If
    End If
    
    ' Check if there is another page
    temp = IsNextPage(MailData)
    If temp <> "" Then
        NextPage = temp
                
        pageNum = pageNum + 1
        lbl = "Fetching Page " & pageNum + 1 & "..."
        
        MailData = ""
        ReadBox = False
        'getNextPage = True
        BatchNumber = 3
        Socket.Action = 2
        Exit Sub
    End If
    
    ' At this point, all the pages have been downloaded
    ' We turn of the timeout timer
    timeout = False
    ' and save the message headers to disk
    frmMailbox.SaveMessages
    
    ' Get the compose page URL
    composeurl = GetComposeURL(MailData)
        
    If chkRefresh.Value = 1 Then
        ' Enable the new mail timer
        refresher = True
    End If
    
    loggedin = True
    frmMailbox.RefreshMessages
    
    ' Have we got the folder information?
    If Not gotFolders Then
        FolderURL = GetFolderURL(MailData)
        If FolderURL <> "" Then
            ' Go get the folders page
            BatchNumber = 4
            Socket.Action = 2
        End If
    Else
        ' Get the compose page
        BatchNumber = 7
        Socket.Action = 2
    End If
    
    'ResetAll
    MailData = ""
                   
    Case 4
        If InStr(1, NewData, "<html>") <> 0 Then
            MailData = NewData
            ReadBox = True
        End If
        
        If ReadBox = True Then
            MailData = MailData & NewData
            If InStr(1, NewData, "</html>") <> 0 Then GoTo FoldersFinished
        End If
        Exit Sub
FoldersFinished:
        gotFolders = True
        Socket.Action = SOCKET_DISCONNECT
        lbl = newmessages
        
        ' Process the folder information
        ProcessFolders (MailData)
        
        frmFolders.RefreshFolders
        
        MailData = ""
        'ResetAll
        
        ' Right here is where we want to start getting the compose page
        ' this page holds two key pieces of information we need for sending messages
        ' One of them is an ID and the other might be some encrypted data.
        BatchNumber = 7
        Socket.Action = 2
    Case 5
        If InStr(1, NewData, "<html>") <> 0 Then
            MailData = MailData & NewData
            ReadBox = True
        End If
        If ReadBox = True Then
            MailData = MailData & NewData
            If InStr(1, NewData, "</html>") Then GoTo MsgDone
        End If
        Exit Sub
MsgDone:
        GotMail = True
        lbl = "Done."
        Socket.Action = SOCKET_DISCONNECT
        fnum = FreeFile
            
        ' Extract the message body and any junk we don't want
        MailData = ProcessMessage(MailData)
        
        ' Indicate the message is now on disk
        Messages(msgIDX).cached = True
        ' Indicate the message is not new anymore, if it was
        Messages(msgIDX).newmail = False
        ChangeTip newmessages
        
        ' Save the message to disk
        If Dir(App.Path & "\HMmsg" & Messages(msgIDX).index & ".html") <> "" Then Kill (App.Path & "\HMmsg" & Messages(msgIDX).index & ".html")
        Open App.Path & "\HMmsg" & Messages(msgIDX).index & ".html" For Binary Access Write As #fnum 'Len = Len(MailData)
        Put #fnum, , MailData
        Close #fnum
                
        ' Load up the message in the browser window
        frmMailbox.htmlBody.Navigate App.Path & "\HMmsg" & Messages(msgIDX).index & ".html"
        
        ResetAll
    Case 6
        ' SEND A MESSAGE
        If InStr(1, NewData, "<html>") <> 0 Then
            MailData = MailData & NewData
            ReadBox = True
        End If
        If ReadBox = True Then
            MailData = MailData & NewData
            If InStr(1, NewData, "</html>") Then GoTo SendDone
        End If
        Exit Sub
SendDone:
        frmhotmail.refresher = True
        GotMail = True
        Socket.Action = SOCKET_DISCONNECT
        If InStr(1, MailData, "Your message has been <b>instantly") <> 0 Then
            PlayWave App.Path & "\msgsent.wav"
            ShowNotify "Your Message Has Been Sent", picIcon(0).Picture
        Else
            ShowNotify "Message Could not be sent", picIcon(5).Picture
        End If
        
        ChangeTip newmessages
        
        ResetAll
    Case 7
        ' Getting the compose page
        If InStr(1, NewData, "<html>") <> 0 Then
            MailData = MailData & NewData
            ReadBox = True
        End If
        If ReadBox = True Then
            MailData = MailData & NewData
            If InStr(1, NewData, "</html>") Then GoTo ComposePageDone
        End If
        Exit Sub
ComposePageDone:
        lbl = newmessages
        GotMail = True
        
        blinker = False
        iconnum = 0
        
        ' I consider this the end of the checking procedure
        ' So, this is where we tell the user what's up
        If newmessages <> "You have no new messages" Then
            ShowNotify newmessages, picIcon(2).Picture
            
            PlayWave App.Path & "\newmail.wav"
            ChangeIcon picIcon(2).Picture, picNotify
            'ShowTip lbl
        Else
            ChangeIcon picIcon(1).Picture, picNotify
            ShowTip "No New Messages"
        End If
        
        Socket.Action = SOCKET_DISCONNECT
        
        fnum = FreeFile
        ' We're going to save it for reference
        If Dir(App.Path & "\compose.html") <> "" Then Kill (App.Path & "\compose.html")
        Open App.Path & "\compose.html" For Binary Access Write As #fnum 'Len = Len(MailData)
        Put #fnum, , MailData
        Close #fnum
        
        ' Now, we PROCESS the compose page.  More details in the
        ' ProcessComposePage function
        ProcessComposePage MailData
        
        ResetAll
    Case 8
        ' DELETE A MESSAGE
        ' Just look for the "<title>Inbox</title>"
        If InStr(1, NewData, "<title>Inbox</title>") <> 0 Then
            ' Don't bother refreshing all the pages.  Just delete the msg.
            Socket.Action = SOCKET_DISCONNECT
            frmMailbox.DeleteMessage frmMailbox.dlItem.index
            
            ShowTip "Message Was Deleted"
            GotMail = True
            GoTo DeleteDone
        End If
        If InStr(1, NewData, "</html>") <> 0 Then
            ShowTip "Message Could not be deleted"
            GotMail = True
            GoTo DeleteDone
        End If
        Exit Sub
DeleteDone:
        ChangeTip newmessages
        ResetAll
End Select
End Sub

Private Sub ConnectToHotmail()
    StrLogin = cboLogin.text
    StrPass = txtpass
    ' Only connect if there is a network connection
    If IsNetConnectOnline = True Then
        ResetIsOnline
        loggedin = False
        frmMailbox.RefreshMessages
        newmessages = ""
        GotMail = False
        timeout = True
        
        ChangeIcon picIcon(0).Picture, picNotify
        lbl.Caption = "1. Connecting to Hotmail..."
        Socket.HostName = "lc5.law5.hotmail.passport.com"
        Socket.Action = 2
    End If
End Sub

Private Sub ResetAll()
On Error Resume Next
Socket.Action = SOCKET_DISCONNECT
BatchNumber = 0

pageNum = 0
'For i = 0 To 5
'Cookies(i) = ""
'Next
CurrentCookie = 0
ReadBox = False
MailData = ""
'NewHost = ""
'NewUrl = ""
End Sub

Private Sub timeout_Timer()
    If timeout.Tag = "" Then timeout.Tag = "0"
    
    If chkTimeout.Value = 1 Then
        timeout.Tag = CLng(timeout.Tag) + 1
        If CLng(timeout.Tag) >= CLng(txtTimeout) Then
            pageNum = 0
            NextPage = NewUrl
            ResetAll
            Socket.Action = SOCKET_DISCONNECT
            If IsNetConnectOnline Then
                ' Still connected.  Maybe just "net congestion" :-)
                ShowTip "Connection Attempt Timed Out. Retrying..."
                loggedin = False
                Socket.Action = SOCKET_DISCONNECT
                ConnectToHotmail
            Else
                ' The computer is no longer connected to the internet
                ShowTip "Connection Attempt Timed Out. Network Disconnected."
                loggedin = False
                ChangeTip "Hotmail Checker NetDetect"
                ChangeIcon picIcon(6), picNotify
                ' Go into NetDetect mode
                nettimer = True
            End If
        End If
    End If
End Sub

Private Sub txtInterval_Change()
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Interval", txtInterval
End Sub

Private Sub txtTimeout_Change()
    REGSaveSetting vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "TimeoutValue", txtTimeout
End Sub

Private Sub RefreshAccounts()
    Dim i As Integer
    cboLogin.Clear
    For i = 0 To AccCount - 1
        If RTrim(Accounts(i).username) <> "" Then
            cboLogin.AddItem RTrim(Accounts(i).loginname)
        End If
    Next
End Sub
