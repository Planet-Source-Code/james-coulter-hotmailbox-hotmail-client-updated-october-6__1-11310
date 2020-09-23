VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFolders 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folders"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "frmFolders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList images 
      Left            =   5280
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFolders.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFolders 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "images"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Messages"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "New"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Size"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.Label lblFolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Folders"
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
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4280
      X2              =   120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   0
      Picture         =   "frmFolders.frx":0896
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub RefreshFolders()
    Dim tItem As ListItem
    Dim i As Integer
    
    lvwFolders.ListItems.Clear
    For i = 0 To FolderCount - 1
        With Folders(i)
            Set tItem = lvwFolders.ListItems.Add(, .id, .fname, , 1)
        End With
    Next
End Sub
