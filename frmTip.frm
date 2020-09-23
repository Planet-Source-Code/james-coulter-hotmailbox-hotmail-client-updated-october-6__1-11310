VERSION 5.00
Begin VB.Form frmTip 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   495
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer closer 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   3240
      Top             =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hotmail Checker:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   40
      Width           =   1410
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   0
      Top             =   45
      Width           =   480
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub closer_Timer()
    Unload Me
End Sub

Private Sub Form_Load()
    closer = True
End Sub
