VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddresses 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Address Book"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmAddresses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   4560
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
            Picture         =   "frmAddresses.frx":0CCA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtEMail 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin MSComctlLib.ListView lvwAdds 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "E-Mail"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.Label lblFolder 
      BackStyle       =   0  'Transparent
      Caption         =   "Address Book"
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
      TabIndex        =   11
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
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   840
      Y2              =   2640
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   360
      X2              =   120
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4320
      X2              =   120
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail:"
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
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   4320
      X2              =   2160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Contact"
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
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   7080
      Left            =   0
      Picture         =   "frmAddresses.frx":19A6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmAddresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public selAddresses As String
Private Sub cmdAdd_Click()
    AddAddress txtName, txtEMail
    SaveAddresses
    RefreshAddresses
    txtName = ""
    txtEMail = ""
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    If Not (lvwAdds.SelectedItem Is Nothing) Then
        Addresses(lvwAdds.SelectedItem.index - 1).fullname = ""
        SaveAddresses
        LoadAddresses
        RefreshAddresses
    End If
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
        
    selAddresses = ""
    For i = 1 To lvwAdds.ListItems.count
        If lvwAdds.ListItems(i).Selected Then
            If selAddresses = "" Then
                selAddresses = lvwAdds.ListItems(i).SubItems(1)
            Else
                selAddresses = selAddresses & "," & lvwAdds.ListItems(i).SubItems(1)
            End If
        End If
    Next
    
    Unload Me
End Sub

Private Sub Form_Load()
    cmdOK.Visible = False
    LoadAddresses
    RefreshAddresses
End Sub

Private Sub AddAddress(ByVal fullname As String, ByVal email As String)
    ReDim Preserve Addresses(addcount) As HotmailAddress
    With Addresses(addcount)
        .fullname = fullname
        .email = email
    End With
    addcount = addcount + 1
End Sub

Private Sub LoadAddresses()
    Dim filename As String
    Dim tAdd As HotmailAddress
    Dim fNum As Integer
    Dim recNum As Long
    filename = App.Path & "\Addresses.idx"
    If Dir(filename) <> "" Then
        addcount = 0
        ReDim Addresses(addcount) As HotmailAddress
        recNum = 1
        ' Get the next availble file number
        fNum = FreeFile
        Open filename For Random Access Read As #fNum Len = Len(tAdd)
        Do Until EOF(fNum)
            ' Seek to the current record number
            Seek #fNum, recNum
            ' Get the address record from the file
            Get #fNum, , tAdd
            recNum = recNum + 1
            ' Make sure it's not a dud
            If tAdd.fullname <> "" And Left(tAdd.fullname, 1) <> Chr(0) Then
                ' Copy the information into the address array
                AddAddress tAdd.fullname, tAdd.email
            End If
            
            If recNum * Len(tAdd) > LOF(fNum) Then
                Exit Do
            End If
        Loop
        Close #fNum
    End If
End Sub
Private Sub SaveAddresses()
    Dim filename As String
    Dim tAdd As HotmailAddress
    Dim fNum As Integer
    filename = App.Path & "\Addresses.idx"
    If Dir(filename) <> "" Then Kill (filename)
        ' Get the next availble file number
        fNum = FreeFile
        Open filename For Random Access Write As #fNum Len = Len(tAdd)
        For i = 0 To addcount - 1
            ' Put the address to the file
            If RTrim(Addresses(i).fullname) <> "" Then
                Put #fNum, , Addresses(i)
            End If
        Next
        Close #fNum
End Sub
Private Sub RefreshAddresses()
    Dim tItem As ListItem
    
    lvwAdds.ListItems.Clear
    For i = 0 To addcount - 1
        With Addresses(i)
            Set tItem = lvwAdds.ListItems.Add(, , RTrim(.fullname), , 1)
            tItem.SubItems(1) = RTrim(.email)
        End With
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveAddresses
End Sub
