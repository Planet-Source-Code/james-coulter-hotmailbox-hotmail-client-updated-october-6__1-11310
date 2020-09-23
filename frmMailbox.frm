VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmMailbox 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Hotmail Messages"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10380
   Icon            =   "frmMailbox.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFolders 
      Caption         =   "Folders"
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
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddresses 
      Caption         =   "Addresses"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   1095
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
      Left            =   3720
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdCompose 
      Caption         =   "Compose"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.Timer blinker 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9840
      Top             =   120
   End
   Begin MSComctlLib.ImageList images 
      Left            =   6720
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailbox.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailbox.frx":0D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailbox.frx":1172
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailbox.frx":1A4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailbox.frx":232A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailbox.frx":277E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailbox.frx":2BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailbox.frx":34AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMailbox.frx":384A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser htmlBody 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   10095
      ExtentX         =   17806
      ExtentY         =   5741
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.ListView lvwMessages 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "images"
      SmallIcons      =   "images"
      ColHdrIcons     =   "images"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "From"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   3
      Top             =   120
      Width           =   5655
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4280
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lblFolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Inbox"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   0
      Picture         =   "frmMailbox.frx":3BE6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmMailbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim linkURL As String

' Record Type for storing the Hotmail message headers
Private Type HotmailMsgDat
    ' The subject
    subject As String * 128
    ' The sender's/destination's name
    sender As String * 128
    ' The sender's/destination's address
    email As String * 128
    ' The date
    indate As String * 64
    ' Is this a new message
    newmail As Boolean
    ' The URL for retrieveing the message
    msgURL As String * 256
    ' Is there an attachment for this message
    attach As Boolean
    ' Size of the message in KB
    size As Long
    index As Integer
    ' Is the message on the drive?
    cached As Boolean
End Type

Dim msgs() As HotmailMsgDat
Dim mcount As Integer

Dim busy As Boolean
' The message that is currently being downloaded (for display purposes)
Public dlItem As ListItem

Public Sub LoadMessages()
    Dim filename As String
    Dim tMsg As HotmailMsgDat
    Dim fnum As Integer
    Dim recNum As Long
    filename = App.Path & "\HMsgs.idx"
    If Dir(filename) <> "" Then
        mcount = 0
        MsgCount = 0
        recNum = 1
        ReDim msgs(mcount) As HotmailMsgDat
        ' Get the next availble file number
        fnum = FreeFile
        Open filename For Random Access Read As #fnum Len = Len(tMsg)
        Do Until EOF(fnum)
            ReDim Preserve msgs(mcount) As HotmailMsgDat
            ' Seek to the current record number
            Seek #fnum, recNum
            ' Get the message header record from the file
            Get #fnum, , msgs(mcount)
            recNum = recNum + 1
            ' Make sure it's not a dud
            If msgs(mcount).msgURL <> "" And Left(msgs(mcount).msgURL, 1) <> Chr(0) Then
                ' Copy the information into the message array
                ReDim Preserve Messages(MsgCount) As HotmailMsg
                With Messages(MsgCount)
                    .attach = msgs(mcount).attach
                    .cached = msgs(mcount).cached
                    .email = RTrim(msgs(mcount).email)
                    .indate = RTrim(msgs(mcount).indate)
                    .index = msgs(mcount).index
                    .msgURL = RTrim(msgs(mcount).msgURL)
                    .newmail = msgs(mcount).newmail
                    .sender = RTrim(msgs(mcount).sender)
                    .size = msgs(mcount).size
                    .subject = RTrim(msgs(mcount).subject)
                End With
                MsgCount = MsgCount + 1
            End If
            
            If recNum * Len(HotmailMsgDat) > LOF(fnum) Then
                Exit Do
            End If
        Loop
        Close #fnum
    End If
End Sub
Public Sub SaveMessages()
    Dim fnum As Integer
    Dim recNum As Long, i As Integer
    Dim filename As String
    Dim tMsg As HotmailMsgDat
    filename = App.Path & "\HMsgs.idx"
    ' Get the next availble file number
    fnum = FreeFile
    If Dir(filename) <> "" Then Kill (filename)
    Open filename For Random Access Write As #fnum Len = Len(tMsg)
    For i = 0 To MsgCount - 1
        If Messages(i).msgURL <> "" And Messages(i).isonline Then
            ' Copy the message details to the record
            With Messages(i)
                tMsg.attach = .attach
                tMsg.cached = .cached
                tMsg.email = .email
                tMsg.indate = .indate
                tMsg.index = .index
                tMsg.msgURL = .msgURL
                tMsg.newmail = .newmail
                tMsg.sender = .sender
                tMsg.size = .size
                tMsg.subject = .subject
            End With
            ' Dump the record to the file
            Put #fnum, , tMsg
        End If
    Next
    Close #fnum
End Sub

Public Sub RefreshMessages()
    Dim i As Integer, j As Integer, tItem As ListItem
    Dim iconnum As Integer
    Dim newmail As Integer
    
    lvwMessages.ListItems.Clear
    For i = MsgCount - 1 To 0 Step -1
        With Messages(i)
            If .isonline Then
ShowMessage:
                If .msgURL <> "" Then
                    If .newmail Then
                        iconnum = 1
                        newmail = newmail + 1
                    Else
                        If .cached Then
                            iconnum = 2
                        Else
                            iconnum = 6
                        End If
                    End If
                    
                    ' If the message's status is yet to be determined,
                    ' we'll show a different icon.
                    If Not .isonline Then
                        ' We'll leave the cached messages as they are
                        If iconnum <> 2 Then
                            iconnum = 7
                        End If
                    End If

                    Set tItem = lvwMessages.ListItems.Add(, .msgURL, .sender, , iconnum)
                    tItem.SubItems(1) = .subject
                    tItem.SubItems(2) = .indate
                    tItem.SubItems(3) = .size & "k"
                    
                    tItem.Key = .msgURL
                    tItem.Tag = .email
                End If
            Else
                ' Show the messages if we aren't logged in yet, even if
                ' they might not be available still.
                If Not loggedin Then
                    GoTo ShowMessage
                End If
            End If
        End With
    Next
    
    If newmail > 0 Then
        ' Don't show new messages if we're not logged in
        If loggedin Then
            If newmail > 1 Then
                lblStatus = "You have " & newmail & " new messages of " & lvwMessages.ListItems.count & " messages."
            Else
                lblStatus = "You have " & newmail & " new message of " & lvwMessages.ListItems.count & " messages."
            End If
        Else
            lblStatus = "You have no new messages of " & lvwMessages.ListItems.count & " messages."
        End If
    Else
        lblStatus = "You have no new messages of " & lvwMessages.ListItems.count & " messages."
    End If
End Sub

Private Sub blinker_Timer()
    On Error Resume Next
    If Not (dlItem Is Nothing) Then
        With dlItem
            If .SmallIcon = 3 Then
                .SmallIcon = 4
            ElseIf .SmallIcon = 4 Then
                .SmallIcon = 3
            End If
        End With
    End If
End Sub

Private Sub cmdAddresses_Click()
    frmAddresses.cmdCancel.Visible = False
    frmAddresses.Show
End Sub

Private Sub cmdCompose_Click()
    frmCompose.Show
End Sub

Private Sub cmdDelete_Click()
    If Not (lvwMessages.SelectedItem Is Nothing) Then
        If GotMail Then
            Dim temp As String
            Dim pos1 As Integer, pos2 As Integer
            BatchNumber = 8
            With lvwMessages.SelectedItem
                pos1 = InStr(1, .Key, "MSG")
                If pos1 <> 0 Then
                    pos2 = InStr(pos1, .Key, "&")
                    If pos2 <> 0 Then
                        GotMail = False
                        frmhotmail.refresher = False
                        Set dlItem = lvwMessages.SelectedItem
                        msgURL = Mid(.Key, pos1, pos2 - pos1)
                    
                        frmhotmail.Socket.Action = SOCKET_DISCONNECT
                        frmhotmail.Socket.Action = 2
                    End If
                End If
            End With
        End If
    End If
End Sub

Private Sub cmdFolders_Click()
    frmFolders.Show
End Sub

Private Sub Form_Load()
    RefreshMessages
    htmlBody.Navigate "http://www.tugboatharbor.com/hmchecker.html"
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        Image1.Height = Height
        lvwMessages.Width = (Width - lvwMessages.Left) - 200
        htmlBody.Width = (Width - htmlBody.Left) - 200
        htmlBody.Height = (Height - htmlBody.Top) - 600
    End If
End Sub

Private Sub htmlBody_BeforeNavigate2(ByVal pDisp As Object, url As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    If InStr(1, url, "compose") Then
        Dim temp As String
        Cancel = True
        
        ' Ah ha!  Here, we can use our Send function.  This URL
        ' is a Hotmail Compose URL.  That means, this is an E-Mail address
        temp = ProcessMailLink(linkURL)
        If temp <> "" Then
            Dim frmNewMail As New frmCompose
            frmNewMail.txtTo = temp
            frmNewMail.Show
            frmNewMail.txtSubject.SetFocus
        End If
    End If
End Sub

Private Sub htmlBody_NavigateComplete2(ByVal pDisp As Object, url As Variant)
    If Not (dlItem Is Nothing) Then
        lvwMessages.Enabled = True
        blinker = False
        On Error Resume Next
        dlItem.SmallIcon = 2
        Set dlItem = Nothing
        busy = False
        frmhotmail.refresher = True
    End If
End Sub

Private Sub htmlBody_StatusTextChange(ByVal text As String)
    linkURL = text
End Sub

Private Sub lvwMessages_Click()
    If Not (lvwMessages.SelectedItem Is Nothing) Then
        With lvwMessages.SelectedItem
            ' If the message is already on the hard drive, display it
            If Messages(MsgCount - .index).cached Then
                msgIDX = MsgCount - .index
                htmlBody.Navigate App.Path & "\HMmsg" & Messages(msgIDX).index & ".html"
            Else
            ' Otherwise, if we have already received the inbox, download the message
                If GotMail Then
                    ' If we aren't already downloading a message
                    'If Not busy Then
                        If Not (dlItem Is Nothing) Then
                            On Error Resume Next
                            If Messages(MsgCount - dlItem.index).newmail Then
                                dlItem.SmallIcon = 1
                            Else
                                If Messages(MsgCount - dlItem.index).cached Then
                                    dlItem.SmallIcon = 2
                                Else
                                    dlItem.SmallIcon = 6
                                End If
                            End If
                        End If
                        frmhotmail.refresher = False
                        msgIDX = MsgCount - .index
                        Set dlItem = lvwMessages.SelectedItem
                        busy = True
                        'lvwMessages.Enabled = False
                        .SmallIcon = 3
                        blinker = True
                        
                        GetMessage .Key
                    'End If
                End If
            End If
        End With
    End If
End Sub

Public Sub DeleteMessage(ByVal index As Integer)
    Set dlItem = Nothing
    frmhotmail.refresher = True
    lvwMessages.ListItems.Remove index
    With Messages(MsgCount - index)
        ' Delete the cached file
        If Dir(App.Path & "\HMmsg" & .index & ".html") <> "" Then
            Kill App.Path & "\HMmsg" & .index & ".html"
        End If
        .msgURL = ""
        .isonline = False
        .index = 0
        .subject = ""
        .cached = False
    End With
    ResizeMessageArray
    SaveMessages
    RefreshMessages
End Sub

' This function extracts a destination E-Mail address from a compose link
Private Function ProcessMailLink(ByVal text As String) As String
    Dim c1 As Integer, c2 As Integer
    Dim temp As String
    
    ' Find the "to=" part
    c1 = InStr(1, text, "&to=")
    If c1 <> 0 Then
        c2 = InStr(c1 + 1, text, "&")
        If c2 <> 0 Then
            c1 = c1 + Len("&to=")
            ' Extract the address
            temp = Mid(text, c1, c2 - c1)
            ' Make sure it is an E-Mail address
            If InStr(1, temp, "@") <> 0 Then
                ProcessMailLink = temp
            End If
        End If
    End If
End Function

