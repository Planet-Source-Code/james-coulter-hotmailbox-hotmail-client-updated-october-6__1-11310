Attribute VB_Name = "SystemTray"
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA = 48


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

'Make your own constant, e.g.:
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205


Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public trayIconNum As Integer
Dim Tic As NOTIFYICONDATA
' Tray icon stuff
Public Sub CreateIcon(picNotify As PictureBox, Optional ByVal ToolTip As String)
    Tic.cbSize = Len(Tic)
    Tic.hwnd = picNotify.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = picNotify.Picture
    Tic.szTip = ToolTip & Chr$(0)
    erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub
Public Sub DeleteIcon(picNotify As PictureBox)
    Tic.cbSize = Len(Tic)
    Tic.hwnd = picNotify.hwnd
    Tic.uID = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Public Sub ShowTip(ByVal text As String)
    'Dim tip As New frmTip
    Dim WindowRect As RECT
    
    Unload frmTip
    
    With frmTip
        .Hide
        .lblTip = text
        .Width = (.lblTip.Left + .lblTip.Width) + 120
        .Height = (.lblTip.Height + .lblTip.Top) + 40
        .Shape1.Height = .Height
        .Shape1.Width = .Width
        
        SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0
        
        .Left = ((WindowRect.Right * Screen.TwipsPerPixelX) - .Width)
        .Top = ((WindowRect.Bottom * Screen.TwipsPerPixelY) - .Height)
        
        If ShowDlgs Then
            .Show
            .SetFocus
            If frmhotmail.chkNoDlgFocus.Value = 0 Then
                pSetForegroundWindow .hwnd
            End If
        End If
    End With
End Sub

Public Sub ShowNotify(ByVal text As String, pic As IPictureDisp)
    Dim note As New frmNotify
    Dim WindowRect As RECT
    SystemParametersInfo SPI_GETWORKAREA, 0, WindowRect, 0

    note.lblStatus = text
    Set note.picIcon.Picture = pic
    note.Width = (note.lblStatus.Left + note.lblStatus.Width) + 400
    note.Left = ((WindowRect.Right * Screen.TwipsPerPixelX) - note.Width)
    note.Top = ((WindowRect.Bottom * Screen.TwipsPerPixelY) - note.Height)
    note.Show
    note.SetFocus
    
    If frmhotmail.chkNoDlgFocus.Value = 0 Then
        pSetForegroundWindow note.hwnd
    End If
End Sub

Public Sub ChangeTip(ByVal text As String)
    Tic.cbSize = Len(Tic)
    'Tic.hwnd = box.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_TIP
    Tic.uCallbackMessage = WM_MOUSEMOVE
    'Set box.Picture = pic
    'Tic.hIcon = box.Picture
    Tic.szTip = text & Chr$(0)
    erg = Shell_NotifyIcon(NIM_MODIFY, Tic)
End Sub

Public Sub ChangeIcon(pic As IPictureDisp, box As PictureBox)
    Tic.cbSize = Len(Tic)
    Tic.hwnd = box.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Set box.Picture = pic
    Tic.hIcon = box.Picture
    If tip <> "" Then
        Tic.szTip = tip & Chr$(0)
    End If
    erg = Shell_NotifyIcon(NIM_MODIFY, Tic)
End Sub
