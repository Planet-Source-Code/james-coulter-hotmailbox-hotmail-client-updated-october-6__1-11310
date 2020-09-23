VERSION 5.00
Begin VB.Form frmNotify 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   540
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer closeTimer 
      Enabled         =   0   'False
      Interval        =   3200
      Left            =   5640
      Top             =   120
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   120
      Picture         =   "frmNotify.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   30
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   520
      Left            =   15
      Top             =   15
      Width           =   6375
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Has Logged On."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   0
      Picture         =   "frmNotify.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub closeTimer_Timer()
    Unload Me
End Sub

Private Sub Form_Load()
    picIcon.BackColor = RGB(47, 103, 144)
    closeTimer = True
End Sub

Private Sub Form_Resize()
    Shape1.Width = Width - 15
    Shape1.Height = Height - 15
End Sub
