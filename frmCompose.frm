VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCompose 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compose New Message"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmCompose.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picNote 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "frmCompose.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   720
      Width           =   480
   End
   Begin VB.TextBox txtBCC 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   2400
      Width           =   2895
   End
   Begin VB.TextBox txtCC 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   2040
      Width           =   2895
   End
   Begin RichTextLib.RichTextBox rtfBody 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3836
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmCompose.frx":0614
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   2760
      Width           =   4335
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   5520
      Width           =   1095
   End
   Begin VB.TextBox txtTo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   2895
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   5520
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompose.frx":0702
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompose.frx":0814
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompose.frx":0926
            Key             =   "Underline"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblFolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Compose"
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
      TabIndex        =   12
      Top             =   120
      Width           =   3975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4280
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Separate multiple addresses with commas.  E-Mail addresses and website URLs will be converted to links automatically."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   720
      TabIndex        =   10
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label lblBCC 
      BackStyle       =   0  'Transparent
      Caption         =   "BCC:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblCC 
      BackStyle       =   0  'Transparent
      Caption         =   "CC:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblTo 
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   0
      Picture         =   "frmCompose.frx":0A38
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmCompose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Bold"
'            rtfBody.SelBold = Not rtfBody.SelBold
'            If rtfBody.SelBold Then
'                Button.Value = tbrPressed
'            Else
'                Button.Value = tbrUnpressed
'            End If
        Case "Italic"
'            rtfBody.SelItalic = Not rtfBody.SelItalic
        Case "Underline"
'            rtfBody.SelUnderline = Not rtfBody.SelUnderline
    End Select
End Sub

Private Sub cmdSend_Click()
    frmhotmail.refresher = False
    If frmhotmail.chkSendSig.Value = 1 Then
        composeString = MakeSendString(txtTo, txtSubject, rtfBody.text, txtCC, txtBCC, frmhotmail.txtSignature)
    Else
        composeString = MakeSendString(txtTo, txtSubject, rtfBody.text, txtCC, txtBCC)
    End If
    BatchNumber = 6
    On Error GoTo StopSend
    GotMail = False
    frmhotmail.Socket.Action = 2
    Unload Me
    Exit Sub
StopSend:
End Sub

Private Sub Form_Load()
    picNote.BackColor = RGB(47, 103, 144)
End Sub

Private Sub lblBCC_Click()
    frmAddresses.cmdOK.Visible = True
    frmAddresses.Show 1
    txtBCC = frmAddresses.selAddresses
    rtfBody.SetFocus
End Sub

Private Sub lblCC_Click()
    frmAddresses.cmdOK.Visible = True
    frmAddresses.Show 1
    txtCC = frmAddresses.selAddresses
    txtBCC.SetFocus
End Sub

Private Sub lblTo_Click()
    frmAddresses.cmdOK.Visible = True
    frmAddresses.Show 1
    txtTo = frmAddresses.selAddresses
    txtCC.SetFocus
End Sub

Private Sub txtSubject_Change()
    Caption = txtSubject
End Sub
