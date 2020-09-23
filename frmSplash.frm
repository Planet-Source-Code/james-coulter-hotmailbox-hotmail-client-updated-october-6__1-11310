VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "HotmailBox"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3975
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer splash 
      Interval        =   5000
      Left            =   2160
      Top             =   3720
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3960
      Left            =   15
      Top             =   15
      Width           =   6735
   End
   Begin VB.Label lblRev 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Revision:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "The MSN logo is a registered trademark of Microsoft Corporation."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 2000, Coulter Software.  All rights reserved."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "The original code for this program was written by Polar J."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3960
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hotmail is a registered trademark of Microsoft Corporation.  All rights Reserved."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   3960
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":8030C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Coulter Software"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   3600
      Width           =   2535
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    lblVersion = App.major & "." & App.minor
    lblRev = App.Revision
    
    Load frmhotmail
    If REGGetSetting(vHKEY_LOCAL_MACHINE, "Software\Hotmail Checker\Settings", "Configured", "0") = "1" Then
        Unload Me
    End If
End Sub


Private Sub splash_Timer()
    Unload Me
End Sub
